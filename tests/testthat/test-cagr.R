test_that("cagr works", {
  expect_equal(round(cagr(9000, 13000, 3), 4), 0.1304)
})
