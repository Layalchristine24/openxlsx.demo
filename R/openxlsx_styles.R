# define style style_variables_names
style_variables_names <- openxlsx::createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  textDecoration = "bold",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "medium",
  wrapText = TRUE
)

# create style to wrap text
style_body <- openxlsx::createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  wrapText = TRUE
)

# define style style_unlocked
style_unlocked <- openxlsx::createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  locked = FALSE,
  fgFill = "#B6D7A8",
  border = c("top", "bottom", "left", "right"),
  wrapText = TRUE
)

# define style style_locked
style_locked <- openxlsx::createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  locked = TRUE,
  fgFill = "#5b5b5b",
  wrapText = TRUE
)

# define style gray
style_gray <- openxlsx::createStyle(
  bgFill = "#AAAAAA"
)

# define style blue
style_blue <- openxlsx::createStyle(
  bgFill = "#6FA8DC"
)

# define style gray
style_green <- openxlsx::createStyle(
  bgFill = "#00AA00"
)

# define style yellow
style_yellow <- openxlsx::createStyle(
  bgFill = "#CCCC00"
)

# define style red
style_red <- openxlsx::createStyle(
  bgFill = "#CC0000",
  fontColour = "#EEEEEE"
)
