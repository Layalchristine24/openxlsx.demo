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
  unquoted_args <- stringr::str_remove_all(names(args_unlocked_cols), '["]')
  unlocked_cols <- stringr::str_remove_all(names(args_unlocked_cols), '["]') %>%
    purrr::map_dbl(~ match(.x, names(data)))

  # find cells which are NA
  tibble::tibble(rows = seq_len(nrow(data))) %>%
    tidyr::crossing(columns = unlocked_cols) %>%
    dplyr::rowwise() %>%
    dplyr::mutate(
      to_unlock = dplyr::case_when(
        is.na(data[[rows, columns]]) ~ 1L,
        TRUE ~ 0L
      )
    ) %>%
    dplyr::select(rows, columns, to_unlock) %>%
    dplyr::ungroup()
}
