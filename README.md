
<!-- README.md is generated from README.Rmd. Please edit that file -->

# openxlsx.demo

<!-- badges: start -->
<!-- badges: end -->

The goal of openxlsx.demo is to create an excel file from scratch only
using `openxlsx` and adding some features one could be interested in.

## Installation

You can install the development version of openxlsx.demo like so:

``` r
# FILL THIS IN! HOW CAN PEOPLE INSTALL YOUR DEV PACKAGE?
```

## Example

This is a basic example which shows you how to create the demo xlsx
file.

``` r
library(openxlsx.demo)

write_penguins(
  data_penguins = palmerpenguins::penguins,
  data_penguins_raw = palmerpenguins::penguins_raw,
  folder_xlsx_file = tempdir()
)
```
