#' Generate a Metadata Report
#'
#' @param dataset A data.frame containing the main dataset.
#' @param metadata_file Path to Excel file with metadata (2 sheets expected).
#' @param output_file Name of the output file to render to (e.g., "report.html").
#' @export
create_metadata_report <- function(df,output_file,reference_file=NULL) {
  if(!grepl(".html",output_file)){
    if(tools::file_path_sans_ext(output_file)==output_file){
      output_file = paste0(output_file,".html")
    }else{
      stop("output_file must be a .html file")
    }
  }
  template_path <- system.file("rmarkdown/templates/metadata_report/skeleton.Rmd",
                               package = "BoldDHSC")
  if(is.null(reference_file)){
  reference_file <- svDialogs::dlg_open(
      title = "Select Reference File",
      filters = matrix(c("Excel", "*.xlsx;*.xls"), ncol = 2, byrow = TRUE)
    )$res

  if(length(reference_file)==0){
     stop("No file selected")}
  }


  if(!grepl("/",output_file)){
    output_path_choose <- svDialogs::dlg_message("You have not selected a folder for your output, do you wish to select an output folder",type="yesno")$res
    if(output_path_choose =="yes"){
      output_path <- svDialogs::dlg_dir(title = "Select Output location")$res
      output_file = paste0(output_path,"/",output_file)
      }
  }

  output_name = tools::file_path_sans_ext(basename(output_file))

  rmarkdown::render(template_path,
                    output_file = output_file,
                    params = list(
                      metadata_file = reference_file,
                      title = output_name,
                      dataset = df
                    ),
                    envir = new.env(parent = globalenv()))
}

