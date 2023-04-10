#' @examples
#' \dontrun{
#' data <- tribble(
#'   ~a, ~year, ~f, ~d, ~b, ~c,
#'   NA, "2019", "g", NA, "f", NA,
#'   "e", NA, "g", "e", NA, NA,
#'   NA, "2021", "g", "e", "f", NA,
#' )
#' find_cells_to_unlock(data = data, "a", "year", "b", "c")
#' }
