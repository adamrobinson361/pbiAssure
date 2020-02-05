#' Read a Power BI layout file into R as a list.
#
#' @param file Power BI layout file
#' @return list
#' @keywords read, layout
#' @export
#' @examples
#' \dontrun{
#' read_layout("Layout)
#' }

read_layout <- function(file){

  file1 <- readLines(file, warn = FALSE, skipNul = TRUE)

  file1 <- enc2utf8(file1)

  RJSONIO::fromJSON(file1)

}

#' Read a Power BI file into R as a list.
#
#' @param file Power BI file
#' @return list
#' @keywords read, layout
#' @export
#' @examples
#' \dontrun{
#' read_layout("example.pbix")
#' }

read_pbi <- function(file){

  temp_dir <- tempdir()

  unzip_folder_name <- stringi::stri_rand_strings(1, 20)

  unzip_folder_path <- file.path(temp_dir, unzip_folder_name)

  unzip(file, exdir = unzip_folder_path)

  read_layout(file.path(unzip_folder_path, "report/Layout"))

}
