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
  wb <- createWorkbook()

  #--- create worksheet --------------------------------------------------------
  # add a new worksheet to the workbook
  ws_penguins <- addWorksheet(
    wb = wb,
    sheetName = "penguins"
  )

  # add a new worksheet to the workbook
  ws_penguins_raw <- addWorksheet(
    wb = wb,
    sheetName = "penguins_raw"
  )

  #--- define first row --------------------------------------------------------
  # first row where to write the data (set to 2 because we want to write comments)
  first_row <- 2L

  #--- modify data -------------------------------------------------------------
  data_penguins_mod <- prepare_penguins_mod(
    data = data_penguins,
    data_raw = data_penguins_raw
  )

  # View(data_penguins_mod)

  #--- write data --------------------------------------------------------------
  # write the palmerpenguins::penguins data
  writeData(
    wb = wb,
    sheet = ws_penguins,
    x = data_penguins_mod,
    startRow = first_row,
    startCol = 1,
    headerStyle = style_variables_names # add a style directly to the header
  )

  # write palmerpenguins::penguins_raw data
  writeData(
    wb = wb,
    sheet = ws_penguins_raw,
    x = data_penguins_raw,
    startRow = first_row,
    startCol = 1,
    headerStyle = style_variables_names, # add a style directly to the header
    withFilter = TRUE # filter on everywhere
  )

  # openXL(wb)

  # --- set columns width ------------------------------------------------------
  # set all cols but any_comment to a specific width in ws_penguins
  setColWidths(
    wb = wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) != "any_comment"),
    widths = 22 # "auto"for automatic sizing
  )

  # set all cols to a specific width in ws_penguins_raw
  setColWidths(
    wb = wb,
    sheet = ws_penguins_raw,
    cols = seq_len(ncol(data_penguins_raw)),
    widths = 22 # "auto"for automatic sizing
  )

  # openXL(wb)

  # --- wrap text --------------------------------------------------------------
  # add style_body to wrap text in ws_penguins
  # (see option 'wrapText = TRUE' in 'createStyle()')
  addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_body,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = seq_len(ncol(data_penguins_mod)),
    gridExpand = TRUE # apply style to all combinations of rows and cols
  )

  # add style_body to wrap text in ws_penguins_raw
  # (see option 'wrapText = TRUE' in 'createStyle()')
  addStyle(
    wb = wb,
    sheet = ws_penguins_raw,
    style = style_body,
    rows = first_row + seq_len(nrow(data_penguins_raw)),
    cols = seq_len(ncol(data_penguins_raw)),
    gridExpand = TRUE # apply style to all combinations of rows and cols
  )

  # openXL(wb)

  #--- date formatting ---------------------------------------------------------
  # see https://ycphs.github.io/openxlsx/articles/Formatting.html#date-formatting

  addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = date_style,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = which(names(data_penguins_mod) == "date_modification"),
    gridExpand = TRUE
  )

  # openXL(wb)

  #--- add drop-down values to size --------------------------------------------
  # add worksheet "Drop-down values" to the workbook
  ws_drop_down_values <- addWorksheet(
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
  writeData(
    wb = wb,
    sheet = "drop-down-values",
    x = options,
    startCol = 1
  )

  # openXL(wb)

  # add drop-downs
  dataValidation(
    wb = wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    operator = "equal",
    type = "list",
    value = "'drop-down-values'!$A$1:$A$5"
  )
  # openXL(wb)

  #--- add colors for drop-down values to size ---------------------------------
  # see https://ycphs.github.io/openxlsx/articles/Formatting.html#conditional-formatting
  conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "huge", # condition under which to apply the formatting
    style = style_gray
  )

  conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "big", # condition under which to apply the formatting
    style = style_blue
  )

  conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "normal", # condition under which to apply the formatting
    style = style_green
  )

  conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "small", # condition under which to apply the formatting
    style = style_yellow
  )

  conditionalFormatting(wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "size"),
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    type = "contains",
    rule = "tiny", # condition under which to apply the formatting
    style = style_red
  )
  # openXL(wb)

  #--- protect worksheets ------------------------------------------------------
  # protect the worksheet ws_penguins
  protectWorksheet(
    wb = wb,
    sheet = ws_penguins,
    lockAutoFilter = FALSE, # allows filtering
    lockFormattingCells = FALSE # allows formatting cells
  )

  # protect the worksheet ws_penguins_raw
  protectWorksheet(
    wb = wb,
    sheet = ws_penguins_raw,
    lockAutoFilter = FALSE, # allows filtering
    lockFormattingCells = FALSE # allows formatting cells
  )

  # openXL(wb)

  #--- unlock column size (ws_penguins) and Comments (ws_penguins_raw) ---------

  # apply unlocked style to size column
  addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = which(names(data_penguins_mod) == "size"),
    gridExpand = TRUE
  )

  # apply unlocked style to Comments column
  addStyle(
    wb = wb,
    sheet = ws_penguins_raw,
    style = style_unlocked,
    rows = first_row + seq_len(nrow(data_penguins_raw)),
    cols = which(names(data_penguins_raw) == "Comments"),
    gridExpand = TRUE
  )

  # openXL(wb)

  #--- lock column any_comments in ws_penguins ---------------------------------

  # apply locked style to any_comment column
  addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_locked,
    rows = first_row + seq_len(nrow(data_penguins_mod)),
    cols = which(names(data_penguins_mod) == "any_comment"),
    gridExpand = TRUE
  )

  # openXL(wb)

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
  addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = first_row + isna_cases$rows,
    cols = isna_cases$columns,
    gridExpand = FALSE
  )

  # apply unlocked style to cells in 1st row of ws_penguins (for comments)
  addStyle(
    wb = wb,
    sheet = ws_penguins,
    style = style_unlocked,
    rows = 1,
    cols = seq_len(ncol(data_penguins_mod)),
    gridExpand = TRUE
  )

  # apply unlocked style to cells in 1st row of ws_penguins_raw (for comments)
  addStyle(
    wb = wb,
    sheet = ws_penguins_raw,
    style = style_unlocked,
    rows = 1,
    cols = seq_len(ncol(data_penguins_raw)),
    gridExpand = TRUE
  )

  # openXL(wb)

  #--- add filter for several variables ----------------------------------------
  # add filtering possibility
  addFilter(
    wb = wb,
    sheet = ws_penguins,
    rows = first_row,
    cols = c(
      which(names(data_penguins_mod) == "year"),
      which(names(data_penguins_mod) == "species"),
      which(names(data_penguins_mod) == "island")
    )
  )

  # openXL(wb)

  #--- freeze the first row and first column -----------------------------------
  # freeze the first row and the first column in ws_penguins
  freezePane(
    wb = wb,
    sheet = ws_penguins,
    firstActiveRow = first_row + 1L,
    firstActiveCol = 2
  )

  # freeze the first row in meta_ws_name
  freezePane(
    wb = wb,
    sheet = ws_penguins_raw,
    firstActiveRow = first_row + 1L
  )

  # openXL(wb)

  #--- hide sheet --------------------------------------------------------------
  # hide sheet "drop-down-values"
  sheetVisibility(wb)[ws_drop_down_values] <- FALSE

  # openXL(wb)

  #--- Internal Hyperlink ------------------------------------------------------
  # internal hyperlink between any_comments (sheet "penguins") and Comments (sheet
  # "penguins_raw")

  # sheet "penguins" should be linked to "penguins_raw"
  write_hyperlink(
    dataset = data_penguins_mod,
    metadata = data_penguins_raw,
    excel_sheet = "penguins",
    first_row = first_row,
    meta_ws_name = "penguins_raw",
    wb = wb
  )

  # Set column any_comment to a specific width in ws_penguins
  # as the comments length is coming from another worksheet and the formula
  # of the internal hyperlink is short, the wrap_text option does not work.
  # Therefore, we need a larger width for the column any_comment.
  setColWidths(
    wb = wb,
    sheet = ws_penguins,
    cols = which(names(data_penguins_mod) == "any_comment"),
    widths = 60 # "auto"for automatic sizing
  )

  # openXL(wb)

  #--- save the workbook -------------------------------------------------------
  # save the workbook
  saveWorkbook(
    wb = wb,
    file = file.path(
      folder_xlsx_file,
      str_c(Sys.Date(), "penguins.xlsx", sep = "_")
    ),
    overwrite = TRUE
  )
}
