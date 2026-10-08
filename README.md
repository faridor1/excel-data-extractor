# Excel Data Extractor
Tool for extracting Excel data. Especially useful for extracting questions and answers from unstructured Excel questionnaires. (R tool -- easy to translate to Python) 

## Motivation

This project addresses a common problem: questionnaires have already been designed and completed without considering how their data will later be analyzed.

Questions and answers may be scattered across formatted worksheets rather than organized into a consistent, analysis-ready table. Recovering that data manually would require opening each questionnaire, locating its questions and answers, and copying them into a structured dataset.

I developed this script to automate that repetitive work across multiple questionnaire files.

## My Approach: Locating Questions from Answer Cells

The central contribution is a layout-based question-matching heuristic that I designed.

Instead of manually specifying the location of every question, the script first identifies answer cells through Excel's data-validation metadata. It then uses each answer cell as the starting point for a spatial search:

1. Search up to ten cells to the right on the same row.
2. If no candidate is found, search left toward column A.
3. If neither search succeeds, search upward in column B.

The first non-empty candidate whose value does not appear among the extracted answers is selected as the question label.

### Why This Search Strategy?

The idea is to exploit common spatial patterns in questionnaires written in left-to-right languages. Question labels often appear beside or above their corresponding answer fields.

The search covers these common positions, with its order reflecting the questionnaire layouts encountered during development.

The distinctive idea is to combine validation metadata with these layout conventions to recover question–answer relationships without requiring the original designers to anticipate data analysis or manually mapping every field afterward.

The approach is designed to work across many common layouts. Its reliability depends on the worksheet structure: writing direction alone does not guarantee correct matching, so extracted results should be reviewed.

## Extraction Workflow

1. Open each Excel workbook as a ZIP archive and read its worksheet XML.
2. Identify answer cells from standard and x14 validation rules.
3. Retrieve dropdown options from the source worksheet and bounds from standard validation rules.
4. Extract the entered answers.
5. Apply the question-matching heuristic.
6. Combine results across files and export an Excel table.

## Example Input

The screenshot below illustrates the questionnaire layout, including question labels and answer fields.

![Questionnaire example](images/questionnaire_layout.png)

## Extracted Output

The screenshot below shows the corresponding extracted records.

![Extraction result](images/scraper_output.png)

Each extracted record contains:

| Column | Description |
| --- | --- |
| `filename` | Source questionnaire filename |
| `answer_address` | Cell address of the answer field |
| `answers` | Entered answer |
| `answer_options` | Dropdown choices or validation bounds |
| `question` | Question label selected by the heuristic |
| `question_address` | Cell address of the selected question |

## Requirements

R with the following packages:

- `dplyr`
- `tidyxl`
- `xml2`
- `stringr`
- `openxlsx`

Install them with:

```r
install.packages(c("dplyr", "tidyxl", "xml2", "stringr", "openxlsx"))
```

## Configuration and Usage

Before running `excel_scraper.R`, update the following settings to match your files:

| Setting | Current assumption |
| --- | --- |
| Working directory | Replace the placeholder in `setwd()` |
| Input folder | Replace the `input_folder` placeholder |
| Output path | Replace the path in `write.xlsx()` |
| Questionnaire worksheet XML | `xl/worksheets/sheet2.xml` |
| Dropdown source worksheet | `Quelle` |
| Questionnaire filename pattern | `Probandenfragen.*\\.xls[a-z]*$` |
| Subfolder naming pattern | Names ending in eight digits |

The script searches matching subfolders immediately inside the input folder. It processes matching questionnaire files inside those subfolders, rather than files directly in the input folder.

After configuration, run the script in R or RStudio:

```r
source("excel_scraper.R")
```

The consolidated results are written to the configured output location.

## Assumptions and Limitations

This implementation was developed for a specific questionnaire structure.

- **Question matching is a heuristic.** Nearby headings or unrelated labels may be selected in other layouts.
- **The upward search uses column B.** It does not search every column above the answer.
- **Candidate filtering compares values.** A question label matching an extracted answer value will be excluded.
- **Cell lookup takes the first matching address across worksheets.** Results therefore depend on the workbook's cell ordering.
- **Validation target ranges are reduced to their first cell.** The script does not expand every cell in those ranges.
- **Dropdown extraction assumes direct cell-range references to `Quelle`.** Other source formats require adaptation.
- **Standard validation bounds are combined as text.** The script does not evaluate the rules or check whether answers satisfy them.
- **The reader supports OOXML workbooks such as `.xlsx` and `.xlsm`.** Legacy `.xls` files are not supported, even though the filename pattern can select them.

## Repository Contents

- `excel_scraper.R` — original R implementation with configurable paths.
- `README.md` — methodology, configuration, and limitations.
- `images/questionnaire_layout.png` — questionnaire screenshot.
- `images/scraper_output.png` — extracted output screenshot.

The underlying questionnaire workbooks and full output dataset are not included.

## Skills Demonstrated

- Semi-structured data extraction
- Excel XML and validation metadata parsing
- Layout-based question matching
- Data transformation and consolidation in R
- Batch processing of questionnaire files
- Traceability through source filenames and cell addresses
