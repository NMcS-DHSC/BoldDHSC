#' Generate a Metadata Report
#'
#' @param dataset A data.frame containing the main dataset.
#' @param metadata_file Path to Excel file with metadata (2 sheets expected).
#' @param output_file Name of the output file to render to (e.g., "report.html").
#' @export
create_metadata_report <- function(output_file = "metadata_report.html") {
  template_path <- system.file("rmarkdown/templates/metadata_report/skeleton.Rmd",
                               package = "BoldDHSC")
  file_path <- svDialogs::dlg_open(
      title = "Select Reference File",
      filters = matrix(c("Excel", "*.xlsx;*.xls"), ncol = 2, byrow = TRUE)
    )$res

  if(length(file_path)==0){
     stop("No file selected")}

  rmarkdown::render(template_path,
                    output_file = output_file,
                    params = list(
                      metadata_file = file_path
                    ),
                    envir = new.env(parent = globalenv()))
}

