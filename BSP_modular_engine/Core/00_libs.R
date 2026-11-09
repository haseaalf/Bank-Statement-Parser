#' Load the packages the engine needs (same approach as BSA's 0_Load_libraries.R).
#'
#' Packages are looked up first in the engine's own library
#' <engine>/r-library/<R major.minor> (filled by Install_BSP.bat, also works
#' offline), then in the user's normal R library. Missing packages are
#' installed into the engine library.

BSP_ENGINE_VERSION <- "1.1.0"
BSP_PACKAGES <- c("pdftools", "openxlsx", "zip")

engine_library <- function(engine_dir) {
  file.path(engine_dir, "r-library", paste(R.version$major, strsplit(R.version$minor, ".", fixed = TRUE)[[1]][1], sep = "."))
}

install_load <- function(packages, lib_dir = NULL) {
  if (!is.null(lib_dir)) {
    dir.create(lib_dir, recursive = TRUE, showWarnings = FALSE)
    .libPaths(c(lib_dir, .libPaths()))
  }
  missing <- packages[!vapply(packages, requireNamespace, logical(1), quietly = TRUE)]
  if (length(missing) > 0) {
    install.packages(missing, lib = if (is.null(lib_dir)) .libPaths()[1] else lib_dir,
                     repos = "https://cloud.r-project.org", type = "win.binary")
  }
  for (package in packages) suppressPackageStartupMessages(do.call("library", list(package)))
}

install_load(BSP_PACKAGES,
             lib_dir = if (exists("R_script_path")) engine_library(R_script_path) else NULL)
