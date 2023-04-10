#' Prepare the data penguins
#'
#' @param data tibble data_penguins
#' @param data_raw tibble data_penguins_raw
#'
#' @return tibble
#' @export
#'
#' @example inst/examples/ex-prepare_penguins_mod.R
prepare_penguins_mod <- function(data,
                                 data_raw) {
  # for demonstration purposes, add a column "any_comment" and "size"
  data |>
    dplyr::mutate(
      size = character(length = nrow(data)),
      any_comment = character(length = nrow(data)),
      id = data_raw$`Individual ID`,
      date_modification = today()
    ) |>
    # rearrange the columns to have 'year' as the first column
    dplyr::select("year", "id", everything()) # use quotes to select global variables
}
