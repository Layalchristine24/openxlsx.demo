#' Prepare the data penguins
#'
#' @param data tibble
#'
#' @return tibble
#' @export
#'
#' @example inst/examples/ex-prepare_penguins_mod.R
prepare_penguins_mod <- function(data) {
  # for demonstration purposes, add a column "any_comment" and "size"
  data |>
    dplyr::mutate(
      size = NA_character_,
      any_comment = "Please add your comment in this field if you feel something is missing."
    ) |>
    # rearrange the columns to have 'year' as the first column
    dplyr::select("year", tidyselect::everything()) # use quotes to select global variables
}
