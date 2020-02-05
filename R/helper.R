#' Helper function hat converts nulls to NA.
#
#' @param x
#' @return character
#' @keywords helper, null, na
#' @examples
#' \dontrun{
#' if_null_na(x)
#' }

if_null_na <- function(x){ifelse(is.null(x), NA, x)}
