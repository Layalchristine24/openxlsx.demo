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
    mutate(
      # find the metadata rows where id matches `Individual ID`
      list_indices_indicators_to_link = as.integer(map(
        id,
        ~ match(
          .x,
          metadata$`Individual ID`
        )
      )),
      # write the link to make the change of a cell value in penguins_raw reactive
      # in the other sheet
      cell = paste0(
        meta_ws_name, "!",
        # get the capital letter for the excel column corresponding to the column
        # index in the penguins_raw dataset
        LETTERS[which(colnames(metadata) == "Comments")],
        list_indices_indicators_to_link + first_row
      ),
      # add an IF condition to get an empty cell if the resp. Comments value in
      # penguins_raw is empty
      link_rewritten = paste0(
        "=IF(", cell, '="","",', cell, ")"
      )
    )

  # write the hyperlink
  writeFormula(
    wb = wb,
    sheet = excel_sheet,
    x = hyperlink_tib$link_rewritten,
    startCol = which(names(dataset) == "any_comment"),
    startRow = as.integer(first_row) + 1L
  )
}
