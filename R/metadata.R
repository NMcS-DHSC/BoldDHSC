#' Create an Excel data dictionary from a dataframe
#'
#' The function creates an excel summary of the dataframe, including missingness, example values, type of column.
#' It also requires the user to input the description for each column.
#' This can either be done in-line as the function runs, or can be uploaded as part of a reference date sheet
#' The file outputted also contains a sheet for the user to explicitly add value to factor columns.
#'
#' @param df A dataframe, ideally of less than 50 columns
#' @param output_file The filepath of the output xlsx file
#' @returns An xlsx document which is saved in the working directory. The first sheet is a summary of all of the columns in the data, and the second sheet is a lookup for each of the factors.
#' @examples
#' metadata(mtcars,"Example_output.xlsx")
#'
metadata <- function(df,output_file) {

  if(substr(output_file,nchar(output_file)-4,nchar(output_file))!=".xlsx"){
    stop("The file name (output_file) must be a .xlsx file")
  }

  suppressWarnings(check <- if(!file.exists(output_file)){TRUE}else{
    tryCatch({
      wb <- openxlsx::loadWorkbook(output_file)  # Fails if read-only
      TRUE}, error = function(e) FALSE)
  }
  )
  if(!check){
    stop("The excel file: ",output_file," is not able to be edited. Please change file name, or close the file before trying again")
  }

  has_ref_file <- svDialogs::dlg_message(
    "Do you have a reference file (in the set format)?",
    type = "yesno"
  )$res

  ref_data <- NULL  # Initialize

  # 2. If "Yes", let them select the file
  if (has_ref_file == "yes") {
    file_path <- svDialogs::dlg_open(
      title = "Select Reference File",
      filters = matrix(c("Excel", "*.xlsx;*.xls"), ncol = 2, byrow = TRUE)
    )$res

    if (length(file_path) > 0) {  # If a file was selected
      # Read the file
      ref_data <- readxl::read_excel(file_path,sheet = 2)
      ref_data_names <- names(ref_data)
      if(all(ref_data_names!= c("Variable_name","Variable_type","Variable_description"))){
        ref_data_check <- FALSE
        svDialogs::dlg_message("Reference data load failed. File not in correct format. Use format from template in this folder",type="ok")
      }else{
        ref_data_check <- TRUE
        message("Reference file loaded successfully!")}
    } else {
      message("No file selected.")
      ref_data_check <- FALSE
    }
  } else {
    message("Proceeding without a reference file.")
    ref_data_check <- FALSE
  }


  if(ncol(df)>100){
    stop("Your dataframe has more than 100 columns. This function will not take a dataframe with more than 100 columns")
  }else if(ncol(df)>50){
    message("Your dataframe has between 50 and 100 columns. This process will still work, but will look messy")
  }

  convert_table_types <- function(data, ref_table) {
    for (col in ref_table$Variable_name) {
      desired_type <- ref_table$Variable_type[ref_table$Variable_name == col]

      if (col %in% names(data)) {
        tryCatch({
          data[[col]] <- switch(
            desired_type,
            "integer" = as.integer(data[[col]]),
            "numeric" = as.numeric(data[[col]]),
            "character" = as.character(data[[col]]),
            "factor" = as.factor(data[[col]]),
            "logical" = as.logical(data[[col]]),
            "date" = as.Date(data[[col]]),
            data[[col]]  # Default: no conversion if type not recognized
          )
        }, warning = function(w) {
          message(sprintf("Warning converting %s to %s: %s", col, desired_type, w$message))
        }, error = function(e) {
          message(sprintf("Failed to convert %s to %s: %s", col, desired_type, e$message))
        })
      } else {
        message(sprintf("Column %s not found in data", col))
      }
    }
    return(data)
  }

  df <- convert_table_types(df,ref_data)


  suppressWarnings(data_dict <- dplyr::summarise(df, across(everything(), ~ list(
                       type = class(.x),
                       n_missing = sum(is.na(.x)),
                       n_available = sum(!is.na(.x)),
                       n_unique = dplyr::n_distinct(.x),
                       example = if (dplyr::n_distinct(.x) >= 4){
                         paste(sample(unique(na.omit(.x)), 3), collapse = ", ")
                       } else {
                         paste(unique(na.omit(.x)), collapse = ", ")
                       },
                       min = ifelse(is.numeric(.x),min(.x),NA),
                       LQ = ifelse(is.numeric(.x),quantile(.x,0.25),NA),
                       median = ifelse(is.numeric(.x),median(.x),NA),
                       mean = ifelse(is.numeric(.x),mean(.x),NA),
                       UQ = ifelse(is.numeric(.x),quantile(.x,0.75),NA),
                       max = ifelse(is.numeric(.x),max(.x),NA)

                     ))))
  data_dict <- as.data.frame(t(data_dict))
  data_dict <- tibble::rownames_to_column(data_dict)
  names(data_dict) <- c("Column_Name","Column_Type","Missing_Values","Non_Missing_Values","Unique_Values","Example_Values","Min","LQ","Median","Mean","UQ","Max")

  data_dict[["Variable_details"]] <- NA

  if(ref_data_check){
    data_dict <- dplyr::left_join(data_dict,ref_data,by=c("Column_Name"="Variable_name"))
    data_dict <- dplyr::mutate(data_dict,Column_Type=ifelse(is.na(Variable_type),Column_Type,Variable_type))
    data_dict <- dplyr::mutate(data_dict,Variable_details=ifelse(is.na(Variable_description),Variable_details,Variable_description))
    data_dict <- dplyr::select(data_dict,-Variable_description,-Variable_type)
  }




  # Loop through each row
  for (i in 1:nrow(data_dict)) {
    # Show the current row data (optional)
    if(is.na(data_dict[i, "Variable_details"])){
      cat("\nCurrent row:\n")
      print(data_dict[i,1 ])

      # Prompt user for input
      user_input <- readline(prompt = paste0("Variable info for: ", data_dict[i,1], ": "))

      # Store the input
      data_dict[i, "Variable_details"] <- user_input
    }
  }


  # Create factor dictionaries including count and percentage columns
  factor_dicts <- list()
  factor_cols <- names(df)[sapply(df, is.factor)]

  for (col in factor_cols) {
    # Compute frequency table including NAs
    freq <- table(df[[col]], useNA = "ifany")

    # Create a data frame from the frequency table
    factor_df <- data.frame(
      `Factor value` = names(freq),
      Count = as.numeric(freq),
      Percent = round(as.numeric(freq) / sum(freq) * 100, 2),
      Description = rep("", length(freq)),
      check.names = FALSE
    )

    factor_dicts[[col]] <- factor_df
  }

  max_chars <- sapply(data_dict, function(col) {
      max(nchar(gsub("[[:punct:]]", "",as.character(col))), na.rm = TRUE)
  })

  column_widths <- ifelse(max_chars<10,16,max_chars*1.2)
  # Create Excel workbook with formatting
  wb <- openxlsx::createWorkbook()

  names(data_dict) <- c("Variable Name",
                        "Variable Type",
                        "Missing values",
                        "Non Missing",
                        "Unique Values",
                        "Example values",
                        "Minimum",
                        "Lower Quartile",
                        "Median",
                        "Mean",
                        "Upper Quartile",
                        "Maximum",
                        "Variable Description")

  # Add main dictionary sheet (red tab)
  openxlsx::addWorksheet(wb, "Data Dictionary")
  openxlsx::writeData(wb, "Data Dictionary", data_dict,na.string = "")
  openxlsx::setColWidths(wb, "Data Dictionary",widths = column_widths,cols = 1:13)

  # Style definitions
  header_style <- openxlsx::createStyle(
    textDecoration = "bold",
    border = "TopBottomLeftRight",
    borderStyle = "medium"
  )

  body_style <- openxlsx::createStyle(
    border = "TopBottomLeftRight",
    borderStyle = "thin"
  )

  outer_border_style <- openxlsx::createStyle(
    border = "TopBottomLeftRight",
    borderStyle = "thin",
    textDecoration = "bold",
    fgFill = "#dddddd"
  )

  # Add factor sheets with formatting
  # Create a single sheet for all factor dictionaries
  sheet_name <- "Factor Dictionaries"
  openxlsx::addWorksheet(wb, sheet_name)

  # Initialize starting row
  current_row <- 1

  for (col_name in names(factor_dicts)) {
    col_info <- data_dict$Variable_details[data_dict$Column_Name==col_name]

    # Write variable name in current row
    openxlsx::writeData(wb, sheet_name, x = col_name, startRow = current_row)
    openxlsx::writeData(wb, sheet_name, x = col_info, startCol = 2, startRow = current_row)

    # Increment row for the table
    current_row <- current_row + 2

    # Write factor table starting at current_row
    openxlsx::writeData(
      wb,
      sheet_name,
      x = factor_dicts[[col_name]],
      startRow = current_row,
      headerStyle = header_style
    )

    # Apply formatting to the current table
    openxlsx::addStyle(
      wb,
      sheet_name,
      style = outer_border_style,
      rows = current_row:(current_row + nrow(factor_dicts[[col_name]])),
      cols = 1:ncol(factor_dicts[[col_name]]),
      gridExpand = TRUE
    )

    openxlsx::addStyle(
      wb,
      sheet_name,
      style = body_style,
      rows = (current_row + 1):(current_row + nrow(factor_dicts[[col_name]]) ),
      cols = 1:ncol(factor_dicts[[col_name]]),
      gridExpand = TRUE
    )

    # Update current_row to be 3 rows below the bottom of this table
    current_row <- current_row + nrow(factor_dicts[[col_name]]) + 3
  }


  # Createt a dataset which looks at completeness per field
  completeness_df <- data.frame(
    Variable = names(df),
    Complete_Count = sapply(df, function(x) sum(!is.na(x))),
    Total_Count = nrow(df),
    Complete_Percent = round(sapply(df, function(x) mean(!is.na(x))) * 100, 2),
    stringsAsFactors = FALSE
  )

  # Create the plot
  completeness_plot <- ggplot2::ggplot(data = completeness_df, ggplot2::aes(x = Variable, y = Complete_Percent)) +
    ggplot2::geom_bar(stat = "identity", fill = "steelblue") +
    ggplot2::coord_flip() +
    ggplot2::labs(
      title = "Missing Data Percentage by Variable",
      x = "Variable",
      y = "Missing Percentage (%)"
    ) +
    ggplot2::theme_minimal()

  # Save the plot as a PNG file
  ggplot2::ggsave(filename = "completeness_plot.png", plot = completeness_plot, width = 6, height = 4, units = "in")


  # Add a worksheet named "Completeness"
  openxlsx::addWorksheet(wb, sheetName = "Completeness")

  # Insert the image into the worksheet
  openxlsx::insertImage(wb, sheet = "Completeness", file = "completeness_plot.png",
                        startRow = 1, startCol = 1, width = 6, height = 4, units = "in")

  # Save workbook
  openxlsx::saveWorkbook(wb, output_file, overwrite = TRUE)
  message(paste("Dictionary saved to:", output_file))
}
