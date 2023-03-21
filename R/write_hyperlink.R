#' @title Write internal hyperlinks
#'
#' @description
#' Write internal hyperlinks between the meta_ws_name sheet and another one to
#' link the Comments variable to the source variable in meta_ws_name.
#'
#' @param dataset dataset containing the any_comment variable to link with the one in
#' the metadata dataset
#' @param excel_sheet name of the excel sheet where the internal link should be
#' written
#' @param metadata metadata dataset where the source variable Comments is stored
#' @param first_row first row where the data should be written in the sheet
#' @param meta_ws_name name of the sheet containing the metadata dataset
#' @param wb name of excel workbook
#'
#' @return invisible
#' @export
write_hyperlink <- function(dataset,
                            metadata,
                            excel_sheet,
                            first_row,
                            meta_ws_name,
                            wb) {
  hyperlink_tib <- dataset |>
    # tibble::rownames_to_column(var = "rowname") |>
    mutate(
      list_indices_indicators_to_link = as.integer(purrr::map(
        id,
        ~ match(
          .x,
          metadata$`Individual ID`
        )
      )),
      link = purrr::map(
        list_indices_indicators_to_link,
        ~ openxlsx::makeHyperlinkString(
          sheet = meta_ws_name,
          row = .x + first_row,
          col = which(names(metadata) == "Comments"),
          text = metadata$Comments[[.x]]
        )
      ),
      # rewrite the link to make the change of a cell value reactive in the other sheet
      cell = paste0(
        meta_ws_name, "!",
        LETTERS[which(colnames(metadata) == "Comments")],
        list_indices_indicators_to_link + first_row
      ),
      link_rewritten = paste0(
        "=IF(", cell, '="","",', cell, ")"
      )
    )

  # write the hyperlink
  openxlsx::writeFormula(
    wb = wb,
    sheet = excel_sheet,
    x = hyperlink_tib$link_rewritten,
    startCol = which(names(dataset) == "any_comment"),
    startRow = as.integer(first_row) + 1L
  )
}
