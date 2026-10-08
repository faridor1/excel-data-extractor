#==============================================================================#
# SCRAPING
#==============================================================================#

# libraries & working directory 
library(dplyr)
library(tidyxl)
library(xml2)
library(stringr)
library(openxlsx)
setwd("set your working directory...")

#------------------------------------------------------------------------------#
# input folder
input_folder <- "direct to your folder containing your excel questionnaires..."

#------------------------------------------------------------------------------#
# wrapping function
convert_questionnaire <- function(file) {
  #------------------------------------------------------------------------------#

# STEP 1: EXTRACT ANSWERS FROM EXCEL METADATA (XML) IDENTIFIED VIA DROP DOWN OPTIONS
  
  # load xml
  # file <- "your questionnaire.xlsm"
  # my .xml file for the questionnaire was "sheet2"
  td <- tempdir()
  unzip(file, exdir =td)
  sheet_xml <- read_xml(paste0(td,"/xl/worksheets/sheet2.xml"))
  
  #------------------------------------------------------------------------------#
  # load excel sheet with hard-coded answer-options ("mine was "Quelle") 
  # cells <- tidyxl::xlsx_cells("your questionnaire.xlsx")
  cells <- tidyxl::xlsx_cells(file)
  
  # extract sheet "Quelle" 
  quelle_cells <- cells %>%
    filter(sheet == "Quelle") %>%
    mutate(value = coalesce(
      as.character(character),
      as.character(numeric),
      as.character(date),
      as.character(logical)
    ))
  
  #------------------------------------------------------------------------------#
  # get cell values 
  get_value <- function(addr, df = cells) {
    df %>%
      filter(address == addr) %>%
      mutate(value = coalesce(
        as.character(character),
        as.character(numeric),
        as.character(date),
        as.character(logical)
      )) %>%
      pull(value) %>%
      .[1]
  }
  
  # get values from hard-coded answer options
  get_range_values <- function(range_str) {
    
    range_str <- str_replace(range_str, ".*!", "")
    range_str <- str_replace_all(range_str, "\\$", "")
    
    start <- str_extract(range_str, "^[A-Z]+[0-9]+")
    end   <- str_extract(range_str, "[A-Z]+[0-9]+$")
    
    col_start <- str_extract(start, "^[A-Z]+")
    row_start <- as.numeric(str_extract(start, "[0-9]+"))
    
    col_end <- str_extract(end, "^[A-Z]+")
    row_end <- as.numeric(str_extract(end, "[0-9]+"))
    
    cols <- openxlsx::col2int(col_start):openxlsx::col2int(col_end)
    rows <- row_start:row_end
    
    addresses <- expand.grid(cols, rows) %>%
      mutate(address = paste0(openxlsx::int2col(Var1), Var2)) %>%
      pull(address)
    
    values <- sapply(addresses, function(a) get_value(a, quelle_cells))
    
    values <- values[!is.na(values)]
    
    paste(values, collapse = "; ")
  }
  #------------------------------------------------------------------------------#
  # "dataValidation" (xml) 
  x14_nodes <- xml_find_all(sheet_xml, ".//*[local-name()='dataValidation' and namespace-uri()='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main']")
  x14_list <- lapply(x14_nodes, function(node) {
    
    # target - cells
    sqref <- xml_text(xml_find_first(node, ".//*[local-name()='sqref']"))
    addresses <- unlist(strsplit(sqref, " "))
    addresses <- str_extract(addresses, "^[^:]+")
    
    # option sources
    formula <- xml_text(xml_find_first(node, ".//*[local-name()='f']"))
    options <- get_range_values(formula)
    
    data.frame(
      answer_address = addresses,
      answer_options = options,
      stringsAsFactors = FALSE
    )
  })
  
  x14_df <- bind_rows(x14_list)
  
  #------------------------------------------------------------------------------#
  # "normal" datavalidation
  dv_nodes <- xml_find_all(sheet_xml, ".//*[local-name()='dataValidation' and not(namespace-uri()='http://schemas.microsoft.com/office/spreadsheetml/2009/9/main')]")
  
  dv_list <- lapply(dv_nodes, function(node) {
    
    sqref <- xml_attr(node, "sqref")
    addresses <- unlist(strsplit(sqref, " "))
    addresses <- str_extract(addresses, "^[^:]+")
    
    f1 <- xml_text(xml_find_first(node, ".//*[local-name()='formula1']"))
    f2 <- xml_text(xml_find_first(node, ".//*[local-name()='formula2']"))
    
    options <- if (!is.na(f1) & !is.na(f2)) {
      paste0(f1, " - ", f2)
    } else {
      NA
    }
    
    data.frame(
      answer_address = addresses,
      answer_options = options,
      stringsAsFactors = FALSE
    )
  })
  
  dv_df <- bind_rows(dv_list)
  
  #------------------------------------------------------------------------------#
  # combine all answer options based on "datavalidations" 
  all_addresses <- unique(c(x14_df$answer_address, dv_df$answer_address))
  
  cond_format <- data.frame(
    answer_address = all_addresses,
    stringsAsFactors = FALSE
  )
  
  #------------------------------------------------------------------------------#
  # read out actual answers from questionnaire 
  cond_format <- cond_format %>%
    rowwise() %>%
    mutate(
      answers = get_value(answer_address)
    ) %>%
    ungroup()
  
  #------------------------------------------------------------------------------#
  # add the answer options
  
  cond_format <- cond_format %>%
    left_join(x14_df, by = "answer_address") %>%
    rename(x14_options = answer_options) %>%
    left_join(dv_df, by = "answer_address") %>%
    rename(dv_options = answer_options) %>%
    mutate(
      answer_options = coalesce(x14_options, dv_options)
    ) %>%
    select(-x14_options, -dv_options)
  
  #------------------------------------------------------------------------------#

# STEP 2: USE SEARCH ALGORITHM TO FIND THE QUESTIONS THAT MATCH ANSWERS

  # Layout-based heuristic: searches right, then left,
  # then above in column B. Assumes a compatible questionnaire layout.
  get_question <- function(addr) {
    
    col <- str_extract(addr, "^[A-Z]+")
    row <- as.numeric(str_extract(addr, "[0-9]+$"))
    col_num <- openxlsx::col2int(col)
    
    # - 1) to the right side 
    for (i in seq(col_num + 1, col_num + 10)) {
      a <- paste0(openxlsx::int2col(i), row)
      val <- get_value(a)
      
      if (!is.na(val) && val != "" &&
          !val %in% cond_format$answers &&
          sum(cond_format$answers == val, na.rm = TRUE) == 0) {
        return(list(
          question=val,
          question_address = a
        ))
      }
    }
    
    # - 2) to the left side 
    for (i in seq(col_num - 1, 1)) {
      a <- paste0(openxlsx::int2col(i), row)
      val <- get_value(a)
      
      if (!is.na(val) && val != "" &&
          !val %in% cond_format$answers &&
          sum(cond_format$answers == val, na.rm = TRUE) == 0) {
        return(list(
          question=val,
          question_address = a
        ))
      }
    }
    
    # - 3) looking above
    for (r in seq(row - 1, 1)) {
      a <- paste0("B", r)
      val <- get_value(a)
      
      if (!is.na(val) && val != "" &&
          !val %in% cond_format$answers &&
          sum(cond_format$answers == val, na.rm = TRUE) == 0) {
        return(list(
          question=val,
          question_address = a
        ))
      }
    }
    
    return(list(
      question = NA,
      question_address = NA
    ))
  }
  
  cond_format <- cond_format %>%
    rowwise() %>%
    mutate(tmp = list(get_question(answer_address))) %>%
    mutate(
      question = tmp$question,
      question_address = tmp$question_address
    ) %>%
    select(-tmp) %>%
    ungroup()
  
  # add file name 
  cond_format <- cond_format %>%
    mutate(filename=basename(file)) %>%
    select(filename,everything())
  
  #------------------------------------------------------------------------------#
  # final dataframe 
  cond_format
  
  #------------------------------------------------------------------------------#
  # end wrapper function
  return(cond_format)
}

#------------------------------------------------------------------------------#
# find directories 
directories <- list.dirs(
  path = input_folder,
  recursive = FALSE,
  full.names = TRUE
)

# only keep directories with questionnaires 
directories <- directories[
  grepl("[0-9]{8}$", basename(directories))
]

cat("Number of directories:", length(directories), "\n\n")

#------------------------------------------------------------------------------#
# find all questionnaires in given folder
# my questionnaires were called "Probandenfragen. ..." 
# insert your questionnaire name pattern. 
files_df <- lapply(seq_along(directories), function(i) {
  
  dir_path <- directories[i]
  
  files <- list.files(
    path = dir_path,
    pattern = "Probandenfragen.*\\.xls[a-z]*$", 
    full.names = TRUE,
    ignore.case = TRUE
  )
  
  data.frame(
    file = files,
    directory = basename(dir_path),
    dir_index = i,
    dir_total = length(directories),
    stringsAsFactors = FALSE
  )
  
}) %>%
  bind_rows()

#------------------------------------------------------------------------------#
# work through all questionnaires
all_results <- lapply(seq_len(nrow(files_df)), function(i) {
  
  f <- files_df$file[i]
  folder <- files_df$directory[i]
  idx <- files_df$dir_index[i]
  total <- files_df$dir_total[i]
  
  cat(
    "Reading file:",
    basename(f),
    "| from",
    folder,
    "| directory",
    paste0("(", idx, "/", total, ")"),
    "\n"
  )
  
  convert_questionnaire(f)
})

#------------------------------------------------------------------------------#
# bind dataframes 
final_df <- bind_rows(all_results)

#------------------------------------------------------------------------------#
# export output file
write.xlsx(
  final_df,
  "directory where your output should go/OUTPUT.xlsx"
)

