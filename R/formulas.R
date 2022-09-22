#' Compound Annual Growth Rate
#'
#' @param begin starting value
#' @param final ending value
#' @param years number of years
#'
#' @return compound annual growth rate
#' @export
#'
#' @examples
#' cagr(9000, 13000, 3)
cagr <- function(begin, final, years) {
  (final / begin)^(1 / years) - 1
}
