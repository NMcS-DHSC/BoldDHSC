#' Create multiple csvs from a fixed format excel document
#'
#' The function creates multiple csv files from a single excel document
#' It is intended to only be used on the pre-defined data dictionary format and will break if the input file isnt in this form
#'
#'
#' @param excel_path The filepath of the input xlsx file
#' @returns Multiple csv documents which are saved in the working directory. The names of the csvs is a concatenation of the main xlsx file name and the sheet names.
#' @examples
#' split_excel_to_csvs("Example_workbook.xlsx")

split_excel_to_csvs <- function(excel_path) {
  # Get base filename without extension
  base_name <- tools::file_path_sans_ext(basename(excel_path))

  # List all sheet names
  sheets <- readxl::excel_sheets(excel_path)

  # Create CSVs
  for (sheet in sheets) {
    # Read the sheet
    data <- readxl::read_excel(excel_path, sheet = sheet)

    # Sanitize sheet name for filename
    safe_sheet <- gsub("[^a-zA-Z0-9_-]", "_", sheet)

    # Create the new filename
    csv_name <- paste0(base_name, "_", safe_sheet, ".csv")

    # Write to CSV
    write.csv(data, csv_name, row.names = FALSE)

    message("Written: ", csv_name)
  }
}

