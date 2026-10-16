# Developer tool: print the lines (with word x positions) the engine sees in a PDF.
# Usage: Rscript tools/dump_lines.R <pdf> [first_page] [last_page]
args <- commandArgs(trailingOnly = TRUE)
engine <- file.path(dirname(normalizePath(sub("--file=", "", grep("--file=", commandArgs(), value = TRUE)))),
                    "..", "BSP_modular_engine")
for (f in sort(list.files(file.path(engine, "Core"), pattern = "\\.R$", full.names = TRUE))) source(f)

doc <- read_statement_pdf(args[1])
p1 <- if (length(args) > 1) as.integer(args[2]) else 1
p2 <- if (length(args) > 2) as.integer(args[3]) else doc$n_pages
for (id in doc$lines$line_id[doc$lines$page >= p1 & doc$lines$page <= p2]) {
  w <- line_words(doc, id)
  cat(sprintf("p%d y%4.0f | %s\n", w$page[1], min(w$y),
              paste(sprintf("%.0f:%s", w$x, w$text), collapse = " ")))
}
