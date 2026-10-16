#' Load the packages the engine needs, installing missing ones (same approach
#' as BSA's 0_Load_libraries.R). An offline library folder can be used by
#' placing the packages in <engine>/r-library.

BSP_ENGINE_VERSION <- "1.0.0"

install_load <- function(packages, lib_dir = NULL) {
  if (!is.null(lib_dir) && dir.exists(lib_dir)) {
    .libPaths(c(lib_dir, .libPaths()))
  }
  for (package in packages) {
    if (!(package %in% rownames(installed.packages()))) {
      install.packages(package, repos = "https://cloud.r-project.org")
    }
    suppressPackageStartupMessages(do.call("library", list(package)))
  }
}

install_load(c("pdftools", "openxlsx"),
             lib_dir = if (exists("R_script_path")) file.path(R_script_path, "r-library") else NULL)
