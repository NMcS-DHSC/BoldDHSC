#' @keywords internal

read_named_tables <- function(file_path) {
  # Get all sheet names
  sheets <- readxl::excel_sheets(file_path)

  # Return empty list if fewer than 6 sheets
  if (length(sheets) < 6) {
    return(list())
  }

  # Only consider sheets after the 5th
  sheets_to_read <- sheets[6:length(sheets)]

  # Initialize result list
  named_tables <- list()

  # Loop through each relevant sheet
  for (sheet in sheets_to_read) {
    # Read cell B1 to get the table name
    #In the template, cell B1 should be the name of the column
    table_name <- readxl::read_excel(file_path, sheet = sheet, range = "B1", col_names = FALSE)[[1,1]]

    # Skip if table_name is NA or empty and return a warning message
    if (is.na(table_name) || table_name == "") {
      print(paste0(sheet," not read due to missing value in cell B1"))
      next
    }
    # Read the data starting from third row
    #The table starts in the third row, so top two rows are skipped
    data <- readxl::read_excel(
      file_path,
      sheet = sheet,
      skip = 2
    )

    # Remove rows where all data is NA
    cleaned_data <- data[!apply(data, 1, function(x) all(is.na(x))), ]

    # Skip if cleaned data has no rows
    if (nrow(cleaned_data) == 0) next

    # Store in the list with the name from A2
    named_tables[[as.character(table_name)]] <- cleaned_data
  }

  # Return the list, or empty list if nothing was added
  return(named_tables)
}
