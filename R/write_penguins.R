#' @title Write penguins datasets into an xlsx file using openxlsx.
#'
#' @description Write the palmerpenguins::penguins and palmerpenguins::penguins_raw
#' datasets in an excel file and add some features.
#'
#' @details Idea: create an excel file from scratch only using openxlsx and adding some
#' features one could be interested in.
#'
#' For our first worksheet, we will use the palmerpenguins::penguins data set
#' (species, island, bill length,bill depth, flipper length, body mass in grams,
#' sex, year). Ideally, we want to modify it and add some features.
#'
#' For the second sheet, we will use the palmerpenguins::penguins_raw data set which
#' contains all the variables and original names as downloaded (please refer to
#' https://allisonhorst.github.io/palmerpenguins/ for more details).
#'
#' @param data_penguins dataset penguins.
#' @param data_penguins_raw dataset penguins_raw.
#' @param folder_xlsx_file directory of the folder where you want to save the
#' xlsx file.
#'
#' @return invisible
#' @export
#'
#' @example inst/examples/ex-write_penguins.R
write_penguins <- function(data_penguins,
                           data_penguins_raw,
                           folder_xlsx_file) {
  # create a new workbook
  wb <- openxlsx::createWorkbook()

  # add a new worksheet to the workbook
  ws_penguins <- openxlsx::addWorksheet(
    wb = wb,
    sheetName = "penguins"
  )

  # first row where to write the data
  first_row <- 1

  # add a new worksheet to the workbook
  ws_penguins_raw <- openxlsx::addWorksheet(
    wb = wb,
    sheetName = "penguins_raw"
  )

  # write the palmerpenguins::penguins data
  openxlsx::writeData(
    wb = wb,
    sheet = ws_penguins,
    x = data_penguins,
    startRow = first_row,
    startCol = 1
  )

  # write palmerpenguins::penguins_raw data
  openxlsx::writeData(
    wb = wb,
    sheet = ws_penguins_raw,
    x = data_penguins_raw,
    startRow = first_row,
    startCol = 1,
    withFilter = TRUE # filter on everywhere
  )


  # openxlsx::openXL(wb)

  # save the workbook
  openxlsx::saveWorkbook(
    wb = wb,
    file = file.path(
      folder_xlsx_file,
      stringr::str_c(Sys.Date(), "penguins.xlsx", collapse = "_")
    ),
    overwrite = FALSE
  )
}
