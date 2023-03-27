test_that("find_cells_to_unlock finds the right cells to unlock", {
  expect_equal(
    sum(find_cells_to_unlock(
      data = tibble::tribble(
        ~a, ~e, ~f, ~d, ~b, ~c,
        NA, "u", NA, "O", "f", "g",
        "e", NA, "g", "e", "V", "g",
        NA, "t", "g", "e", "f", "B"
      ),
      "a", "b", "c"
    )$to_unlock),
    2
  )
})
