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
  #--- date formatting ---------------------------------------------------------
  # add option for the date formatting
  options(openxlsx.dateFormat = "yyyy/mm/dd")

  #--- create workbook ---------------------------------------------------------
  # create a new workbook
  wb <- openxlsx::createWorkbook()

  #--- create worksheet --------------------------------------------------------
  # add a new worksheet to the workbook
  ws_penguins <- openxlsx::addWorksheet(
    wb = wb,
    sheetName = "penguins"
  )

  # add a new worksheet to the workbook
  ws_penguins_raw <- openxlsx::addWorksheet(
    wb = wb,
    sheetName = "penguins_raw"
  )

  #--- define first row --------------------------------------------------------
  # first row where to write the data
  first_row <- 1

  #--- modify data -------------------------------------------------------------
  # for demonstration purposes, add a column "consistency"
  data_penguins_mod <- data_penguins |>
    dplyr::mutate(
      size = NA_character_,
      date_today = lubridate::today()
    )

  #--- write data --------------------------------------------------------------
  # write the palmerpenguins::penguins data
  openxlsx::writeData(
    wb = wb,
    sheet = ws_penguins,
    x = data_penguins_mod,
    startRow = first_row,
    startCol = 1,
    headerStyle = style_variables_names # add a style directly to the header
  )

  # write palmerpenguins::penguins_raw data
  openxlsx::writeData(
    wb = wb,
    sheet = ws_penguins_raw,
    x = data_penguins_raw,
    startRow = first_row,
    startCol = 1,
    headerStyle = style_variables_names, # add a style directly to the header
    withFilter = TRUE # filter on everywhere
  )

  # openxlsx::openXL(wb)

  # --- set columns width ------------------------------------------------------
  # set all cols to a set width and wrap text in ws_penguins
  openxlsx::setColWidths(
    wb = wb,
    sheet = ws_penguins,
    cols = seq_len(ncol(data_penguins_mod)),
    widths = 20 # "auto"for automatic sizing
  )

  # add style to wrap text
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_wraptext,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = seq_len(ncol(data_penguins_mod)),
    gridExpand = TRUE # apply style to all combinations of rows and cols
  )

  # set all cols to a set width and wrap text in ws_penguins_raw
  openxlsx::setColWidths(
    wb = wb,
    sheet = ws_penguins_raw,
    cols = seq_len(ncol(data_penguins_raw)),
    widths = 27
  )

  # add style to wrap text
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins_raw,
    style = style_wraptext,
    rows = first_row + seq_len(nrow(data_penguins_raw)),
    cols = seq_len(ncol(data_penguins_raw)),
    gridExpand = TRUE # apply style to all combinations of rows and cols
  )

  # openxlsx::openXL(wb)

  #--- add drop-down values to size --------------------------------------------
  # add worksheet "Drop-down values" to the workbook
  openxlsx::addWorksheet(
    wb = wb,
    sheetName = "drop-down-values"
  )

  # add options for the drop-down in a second sheet
  options <- c(
    "huge",
    "big",
    "normal",
    "small",
    "tiny"
  )

  # add drop-down values dataframe to the sheet "Drop-down values"
  openxlsx::writeData(
    wb = wb,
    sheet = "drop-down-values",
    x = options,
    startCol = 1
  )

  # add drop-downs
  openxlsx::dataValidation(
    wb = wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + 1:nrow(data_penguins_mod),
    type = "list",
    value = "'drop-down-values'!$A$1:$A$5"
  )
  # openxlsx::openXL(wb)

  #--- add colors for drop-down values to size ---------------------------------
  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = 2:(nrow(data_penguins_mod) + 1L),
    type = "contains",
    rule = "huge",
    style = openxlsx::createStyle(
      bgFill = "#AAAAAA"
    )
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = 2:(nrow(data_penguins_mod) + 1L),
    type = "contains",
    rule = "big",
    style = openxlsx::createStyle(
      bgFill = "#6FA8DC"
    )
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = 2:(nrow(data_penguins_mod) + 1L),
    type = "contains",
    rule = "normal",
    style = openxlsx::createStyle(
      bgFill = "#00AA00"
    )
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = 2:(nrow(data_penguins_mod) + 1L),
    type = "contains",
    rule = "small",
    style = openxlsx::createStyle(
      bgFill = "#CCCC00"
    )
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = 2:(nrow(data_penguins_mod) + 1L),
    type = "contains",
    rule = "tiny",
    style = openxlsx::createStyle(
      bgFill = "#CC0000",
      fontColour = "#EEEEEE"
    )
  )
  # openxlsx::openXL(wb)


  #--- protect worksheet -------------------------------------------------------
  # protect the worksheet ws_penguins
  openxlsx::protectWorksheet(
    wb = wb,
    sheet = ws_penguins,
    lockAutoFilter = FALSE, # allows filtering
    lockFormattingCells = FALSE # allows formatting cells (Kommentare)
  )

  # openxlsx::openXL(wb)

  #--- unlock specific cells ---------------------------------------------------
  # indices of columns to be unlocked if no value in a cell
  unlocked_cols <- c(
    which(names(data_penguins_mod) == "body_mass_g"),
    which(names(data_penguins_mod) == "sex")
  )

  # indices of the rows to be unlocked if no value for body_mass_g and sex
  indices_isna_body_mass_g <- which(
    is.na(data_penguins_mod$body_mass_g)
  )
  indices_isna_sex <- which(
    is.na(data_penguins_mod$sex)
  )

  # filter all is_na cells
  tib_indices <- tidyr::crossing(
    # only the rows regarding if no value for body_mass_g
    rows = c(indices_isna_body_mass_g, indices_isna_sex),
    # only the columns for the variables Wert, Ampelwert and Begruendung
    columns = unlocked_cols
  ) %>%
    dplyr::rowwise() %>%
    dplyr::mutate(
      isna_cell = dplyr::case_when(
        # if the cell is empty, it means that it can be modified.
        is.na(data_penguins_mod[[rows, columns]]) ~ 1,
        TRUE ~ 0
      )
    )

  # subtable with only na cases
  isna_cases <- tib_indices %>%
    dplyr::filter(isna_cell == 1) %>%
    dplyr::arrange(rows, columns)


  # apply unlocked style to isna_cases cells
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = first_row + isna_cases$rows,
    cols = isna_cases$columns,
    gridExpand = FALSE
  )


  # openxlsx::openXL(wb)

  #--- add filter for several rows ---------------------------------------------
  # add filtering possibility
  openxlsx::addFilter(
    wb = wb,
    sheet = ws_penguins,
    rows = first_row,
    cols = c(
      which(names(data_penguins) == "body_mass_g"),
      which(names(data_penguins) == "sex")
    )
  )

  # openxlsx::openXL(wb)

  #--- freeze the first row and first column -----------------------------------
  # freeze the first row and the first column in ws_penguins
  openxlsx::freezePane(
    wb = wb,
    sheet = ws_penguins,
    firstRow = TRUE,
    firstCol = TRUE
  )

  # freeze the first row in meta_ws_name
  openxlsx::freezePane(
    wb = wb,
    sheet = ws_penguins_raw,
    firstRow = TRUE
  )

  # openxlsx::openXL(wb)

  #--- hide sheet --------------------------------------------------------------
  # hide sheet "drop-down-values"
  openxlsx::sheetVisibility(wb)[3] <- FALSE

  # openxlsx::openXL(wb)

  #--- save the workbook -------------------------------------------------------
  # save the workbook
  openxlsx::saveWorkbook(
    wb = wb,
    file = file.path(
      folder_xlsx_file,
      stringr::str_c(Sys.Date(), "penguins.xlsx", collapse = "_")
    ),
    overwrite = TRUE
  )
}
