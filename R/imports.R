#' Pipe operator
#'
#' See `dplyr::[\%>\%][dplyr::reexports]` for details.
#'
#' @name %>%
#' @rdname pipe
#' @keywords internal
#' @export
#' @importFrom dplyr %>%
#' @usage lhs \%>\% rhs
#' @param lhs A value or the magrittr placeholder.
#' @param rhs A function call using the magrittr semantics.
#' @return The result of calling `rhs(lhs)`.
NULL

#' .data pronoun
#'
#' See `rlang::[.data][rlang::.data]` for details.
#'
#' @name .data
#' @rdname dot-data
#' @keywords internal
#' @export
#' @importFrom rlang .data
#' @inherit rlang::.data description
NULL

#' .env pronoun
#'
#' See `rlang::[.env][rlang::.env]` for details.
#'
#' @name .env
#' @rdname dot-env
#' @keywords internal
#' @export
#' @importFrom rlang .env
#' @inherit rlang::.env description
NULL
