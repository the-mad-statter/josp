#' Generate Resource Paths
#'
#' @param path relative component of path between josp libary and file
#' @param file name of resource
#'
#' @return absolute path to resource file
#'
#' @examples
#' resource("officer", "template_16x9.pptx")
resource <- function(path, file) {
  system.file(sprintf("%s/%s", path, file), package = "josp")
}
