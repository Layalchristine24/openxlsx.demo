#' Find indices of rows and columns of data to be unlocked
#'
#' @param data data containing the cells to be unlock
#' @param ... column names where some cells should be unlocked
#'
#' @return  tibble containing all rows and columns information about
#' which cell should be locked or unlocked
#'
#' @export
#'
#' @example inst/examples/ex-find_cells_to_unlock.R

find_cells_to_unlock <- function(data,
                                 ...) {
  # indices of columns to be unlocked
  args_unlocked_cols <- tibble::lst(...)
  unlocked_cols <- stringr::str_remove_all(names(args_unlocked_cols), '["]') |>
    purrr::map_dbl(~ match(.x, names(data)))

  # find cells which are NA
  tibble::tibble(rows = seq_len(nrow(data))) |>
    tidyr::crossing(columns = unlocked_cols) |>
    mutate(
      to_unlock = purrr::map2_int(rows, columns, function(row, col) {
        if_else(is.na(data[[row, col]]), 1L, 0L)
      })
    ) |>
    select(rows, columns, to_unlock) |>
    ungroup()
}
