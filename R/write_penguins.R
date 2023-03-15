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
  # pkgload::load_all()

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
  # first row where to write the data (set to 2 because we want to write comments)
  first_row <- 2

  #--- modify data -------------------------------------------------------------
  data_penguins_mod <- prepare_penguins_mod(data_penguins)

  # View(data_penguins_mod)

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
    widths = 25 # "auto"for automatic sizing
  )

  # set all cols to a set width and wrap text in ws_penguins_raw
  openxlsx::setColWidths(
    wb = wb,
    sheet = ws_penguins_raw,
    cols = seq_len(ncol(data_penguins_raw)),
    widths = 25 # "auto"for automatic sizing
  )

  # openxlsx::openXL(wb)

  # --- wrap text --------------------------------------------------------------
  # add style_body to wrap text in ws_penguins
  # (see option 'wrapText = TRUE' in 'createStyle()')
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_body,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = seq_len(ncol(data_penguins_mod)),
    gridExpand = TRUE # apply style to all combinations of rows and cols
  )

  # add style_body to wrap text in ws_penguins_raw
  # (see option 'wrapText = TRUE' in 'createStyle()')
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins_raw,
    style = style_body,
    rows = first_row + seq_len(nrow(data_penguins_raw)),
    cols = seq_len(ncol(data_penguins_raw)),
    gridExpand = TRUE # apply style to all combinations of rows and cols
  )

  # openxlsx::openXL(wb)

  #--- add drop-down values to size --------------------------------------------
  # add worksheet "Drop-down values" to the workbook
  ws_drop_down_values <- openxlsx::addWorksheet(
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

  # openxlsx::openXL(wb)

  # add drop-downs
  openxlsx::dataValidation(
    wb = wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    operator = "equal",
    type = "list",
    value = "'drop-down-values'!$A$1:$A$5"
  )
  # openxlsx::openXL(wb)

  #--- add colors for drop-down values to size ---------------------------------
  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "huge", # condition under which to apply the formatting
    style = style_gray
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "big", # condition under which to apply the formatting
    style = style_blue
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "normal", # condition under which to apply the formatting
    style = style_green
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "small", # condition under which to apply the formatting
    style = style_yellow
  )

  openxlsx::conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "tiny", # condition under which to apply the formatting
    style = style_red
  )
  # openxlsx::openXL(wb)

  #--- protect worksheet -------------------------------------------------------
  # protect the worksheet ws_penguins
  openxlsx::protectWorksheet(
    wb = wb,
    sheet = ws_penguins,
    lockAutoFilter = FALSE, # allows filtering
    lockFormattingCells = FALSE # allows formatting cells
  )

  # openxlsx::openXL(wb)

  #--- unlock column size and any_comment --------------------------------------

  # apply unlocked style to size column
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = which(names(data_penguins_mod) == "size"),
    gridExpand = FALSE
  )

  # apply unlocked style to any_comment column
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = which(names(data_penguins_mod) == "any_comment"),
    gridExpand = FALSE
  )
  # openxlsx::openXL(wb)

  #--- unlock specific cells ---------------------------------------------------
  # indices of columns to be unlocked if no value in a cell
  tib_indices <- find_cells_to_unlock(
    data = data_penguins_mod,
    "bill_length_mm", "bill_depth_mm", "flipper_length_mm", "body_mass_g", "sex"
  )

  # subtable with only na cases
  isna_cases <- tib_indices |>
    dplyr::filter(to_unlock == 1) |>
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

  # apply unlocked style to cells in 1st row of ws_penguins (for comments)
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = 1,
    cols = seq_len(ncol(data_penguins_mod)),
    gridExpand = TRUE
  )

  # apply unlocked style to cells in 1st row of ws_penguins_raw (for comments)
  openxlsx::addStyle(
    wb = wb,
    sheet = ws_penguins_raw,
    style = style_unlocked,
    rows = 1,
    cols = seq_len(ncol(data_penguins_raw)),
    gridExpand = TRUE
  )

  # openxlsx::openXL(wb)

  #--- add filter for several variables ----------------------------------------
  # add filtering possibility
  openxlsx::addFilter(
    wb = wb,
    sheet = ws_penguins,
    rows = first_row,
    cols = c(
      which(names(data_penguins_mod) == "year"),
      which(names(data_penguins_mod) == "species"),
      which(names(data_penguins_mod) == "island")
    )
  )

  # openxlsx::openXL(wb)

  #--- freeze the first row and first column -----------------------------------
  # freeze the first row and the first column in ws_penguins
  openxlsx::freezePane(
    wb = wb,
    sheet = ws_penguins,
    firstActiveRow = first_row + 1L,
    firstActiveCol = 2
  )

  # freeze the first row in meta_ws_name
  openxlsx::freezePane(
    wb = wb,
    sheet = ws_penguins_raw,
    firstActiveRow = first_row + 1L
  )

  # openxlsx::openXL(wb)

  #--- hide sheet --------------------------------------------------------------
  # hide sheet "drop-down-values"
  openxlsx::sheetVisibility(wb)[ws_drop_down_values] <- FALSE

  # openxlsx::openXL(wb)

  #--- save the workbook -------------------------------------------------------
  # save the workbook
  openxlsx::saveWorkbook(
    wb = wb,
    file = file.path(
      folder_xlsx_file,
      stringr::str_c(Sys.Date(), "penguins.xlsx", sep = "_")
    ),
    overwrite = TRUE
  )
}
