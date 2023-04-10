# define style style_variables_names
style_variables_names <- createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  textDecoration = "bold",
  border = c("top", "bottom", "left", "right"),
  borderStyle = "medium",
  wrapText = TRUE
)

# create style to wrap text
style_body <- createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  wrapText = TRUE
)

# define style style_unlocked
style_unlocked <- createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  locked = FALSE,
  fgFill = "#d9d2e9",
  border = c("top", "bottom", "left", "right"),
  wrapText = TRUE
)

# define style style_locked
style_locked <- createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  locked = TRUE,
  fgFill = "#f4cccc",
  border = c("top", "bottom", "left", "right"),
  wrapText = TRUE
)

# define style gray
style_gray <- createStyle(
  bgFill = "#AAAAAA"
)

# define style blue
style_blue <- createStyle(
  bgFill = "#6FA8DC"
)

# define style gray
style_green <- createStyle(
  bgFill = "#00AA00"
)

# define style yellow
style_yellow <- createStyle(
  bgFill = "#CCCC00"
)

# define style red
style_red <- createStyle(
  bgFill = "#CC0000",
  fontColour = "#EEEEEE"
)

# define date_style
date_style <- createStyle(numFmt = "dd/mm/yyyy")
