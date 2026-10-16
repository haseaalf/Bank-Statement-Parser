# Developer tool: print an output xlsx. Usage: Rscript tools/show_output.R <xlsx> [max_rows]
args <- commandArgs(trailingOnly = TRUE)
x <- openxlsx::read.xlsx(args[1], colNames = TRUE, check.names = FALSE)
n <- if (length(args) > 1) as.integer(args[2]) else nrow(x)
cat("Sheet:", openxlsx::getSheetNames(args[1]), "| rows:", nrow(x), "\n")
for (i in seq_len(min(n, nrow(x)))) {
  cat(sprintf("%-5s | %-60s | %14.2f | %14.2f | %16.2f\n", x[i, 1],
              gsub("\n", " \\\\n ", x[i, 2]), x[i, 3], x[i, 4], x[i, 5]))
}
