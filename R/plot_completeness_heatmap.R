#' Create a factor defined heatmap of completeness
#'
#' The function creates a summary of missingness in a dataframe, split by a pareticular factor column
#'
#'
#' @param df A dataframe
#' @param factor_col The name of the column which has the factor values in to be compared, given in quotations
#' @returns A heatmap as a ggplot2 object
#' @examples
#' plot_completeness_heatmap(airquality,"Month")
#'
plot_completeness_heatmap <- function(df, factor_col) {
  df %>%
    dplyr::group_by(across(all_of(factor_col))) %>%
    dplyr::summarise(across(everything(), ~mean(!is.na(.)), .names = "complete_{.col}")) %>%
    tidyr::pivot_longer(cols = starts_with("complete_"), names_to = "Variable", values_to = "Completeness") %>%
    dplyr::mutate(Variable = sub("complete_", "", Variable)) %>%
    ggplot2::ggplot(ggplot2::aes(x = Variable, y = .data[[factor_col]], fill = Completeness)) +
    ggplot2::geom_tile(color = "white") +
    ggplot2::scale_fill_gradientn(
      colours = c("darkred","red","orange","yellow", "green"),
      values = scales::rescale(c(0, 0.5, 0.7, 0.9, 1)),
      limits = c(0, 1),
      name = "Completeness"
    )+
    ggplot2::labs(title = paste("Data Completeness by", factor_col),
         x = "Variable",
         y = factor_col,
         fill = "Completeness") +
    ggplot2::theme_minimal() +
    ggplot2::theme(axis.text.x = ggplot2::element_text(angle = 45, hjust = 1))
}

