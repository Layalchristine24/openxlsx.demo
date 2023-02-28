# define style style_variables_names
style_variables_names <- openxlsx::createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  textDecoration = "bold"
)

# create style to wrap text
style_wraptext <- openxlsx::createStyle(
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
  border = c("top", "bottom", "left", "right")
)

# define style style_locked
style_locked <- openxlsx::createStyle(
  fontSize = 12,
  halign = "left",
  valign = "top",
  locked = TRUE,
  fgFill = "#5b5b5b"
)
