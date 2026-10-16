# --------------- Argument Structure -----------------------------------------------------
# 1 : Work folder (contains job_manifest.csv / job_options.csv, receives result files)
# 2 : Folder path of the R engine (this folder)
# 3 : Mode : list_banks | detect | parse
#
# Called from BSP_v1.xlsm:
#   "<R folder>\Rscript.exe" "<engine>\BSP_engine.R" "<work folder>" "<engine>" "<mode>"
#
# Input files (work folder)
#   job_manifest.csv : file_path, bank_id (AUTO or id), variant, password
#   job_options.csv  : key, value  (output_dir, output_mode, on_exists, tolerance, verify_only)
# Output files (work folder)
#   bank_registry.csv, detect_result.csv, run_result.csv, verification_detail.csv,
#   consolidation_detail.csv, engine_log.txt, engine_done.txt

args <- commandArgs(trailingOnly = TRUE)
work_dir      <- args[1]
R_script_path <- args[2]
mode          <- if (length(args) >= 3) tolower(args[3]) else "parse"

done_file <- file.path(work_dir, "engine_done.txt")
if (file.exists(done_file)) invisible(file.remove(done_file))

log_con <- file(file.path(work_dir, "engine_log.txt"), open = "wt", encoding = "UTF-8")
sink(log_con, type = "output")
sink(log_con, type = "message")

finish <- function(status, message = "") {
  writeLines(c(status, message), done_file, useBytes = TRUE)
  sink(type = "message"); sink(type = "output"); close(log_con)
  quit(save = "no", status = 0)
}

main <- function() {
  cat("BSP engine started:", format(Sys.time()), "\n")
  cat("Mode:", mode, "\nWork folder:", work_dir, "\nEngine folder:", R_script_path, "\n")

  # ---- 1. Load core functions and bank parsers (every .R file in Core/ and Banks/)
  for (folder in c("Core", "Banks")) {
    for (f in sort(list.files(file.path(R_script_path, folder), pattern = "\\.R$", full.names = TRUE))) {
      cat("Reading File :", basename(f), "\n")
      source(f, encoding = "UTF-8")
    }
  }
  cat("Engine version", BSP_ENGINE_VERSION, "-", nrow(bank_registry_table()), "bank formats loaded\n")

  write_result_csv(bank_registry_table(), file.path(work_dir, "bank_registry.csv"))
  if (mode == "list_banks") return(invisible())

  # ---- 2. Read job
  manifest <- utils::read.csv(file.path(work_dir, "job_manifest.csv"), stringsAsFactors = FALSE,
                              colClasses = "character", fileEncoding = "UTF-8-BOM")
  for (col in c("bank_id", "variant", "password", "output_dir")) if (!col %in% names(manifest)) manifest[[col]] <- ""
  manifest[is.na(manifest)] <- ""

  opt_file <- file.path(work_dir, "job_options.csv")
  opts <- list(output_dir = work_dir, output_mode = "per_file", on_exists = "suffix",
               tolerance = "0.01", verify_only = "0")
  if (file.exists(opt_file)) {
    o <- tryCatch(utils::read.csv(opt_file, stringsAsFactors = FALSE, colClasses = "character",
                                  fileEncoding = "UTF-8-BOM"),
                  error = function(e) data.frame(key = character(), value = character()))
    if (all(c("key", "value") %in% names(o))) for (i in seq_len(nrow(o))) opts[[o$key[i]]] <- o$value[i]
  }
  tol <- as.numeric(opts$tolerance)
  verify_only <- opts$verify_only %in% c("1", "TRUE", "true")
  cat("Options:", paste(names(opts), unlist(opts), sep = "=", collapse = "; "), "\n")

  if (mode == "detect") run_detect(manifest) else run_parse(manifest, opts, tol, verify_only)
}

# ---- Detect mode: which bank is each file, plus account/period for the preview grid
run_detect <- function(manifest) {
  out <- lapply(seq_len(nrow(manifest)), function(i) {
    m <- manifest[i, ]
    cat("Detect:", m$file_path, "\n")
    base <- data.frame(file_path = m$file_path, file_name = basename(m$file_path),
                       bank_id = "", variant = "", bank_name = "", score = NA, status = "ERROR",
                       message = "", account_no = "", account_name = "", period_start = "",
                       period_end = "", n_pages = NA, stringsAsFactors = FALSE)
    tryCatch({
      doc <- read_statement_pdf(m$file_path, m$password)
      base$n_pages <- doc$n_pages
      r <- resolve_bank(doc, m$bank_id, m$variant)
      base$status <- r$status; base$message <- r$message; base$score <- round(r$score, 2)
      if (!is.null(r$bank)) {
        base$bank_id <- r$bank$id; base$variant <- r$bank$variant; base$bank_name <- r$bank$name
        st <- r$bank$parse(doc)
        base$account_no <- one_line(st$meta$account_no)
        base$account_name <- one_line(st$meta$account_name)
        base$period_start <- fmt_date(st$meta$period_start)
        base$period_end <- fmt_date(st$meta$period_end)
      }
      base
    }, error = function(e) {
      base$message <- conditionMessage(e)
      if (inherits(e, "bsp_error") && e$code == "PDF_PASSWORD") base$status <- "PASSWORD"
      base
    })
  })
  write_result_csv(do.call(rbind, out), file.path(work_dir, "detect_result.csv"))
}

#' Write one output file; returns TRUE or an error message (e.g. file open in Excel)
safe_write <- function(txns, sheet, file) {
  tryCatch({
    dir.create(dirname(file), recursive = TRUE, showWarnings = FALSE)
    write_output_xlsx(txns, sheet, file)
    cat("Written:", file, "\n")
    TRUE
  }, error = function(e) {
    cat("ERROR writing", file, ":", conditionMessage(e), "\n")
    paste0("Gagal menyimpan output (file sedang dibuka atau folder tidak bisa ditulis): ", conditionMessage(e))
  })
}

# ---- Parse mode
run_parse <- function(manifest, opts, tol, verify_only) {
  results <- list(); details <- list(); cons_details <- list()
  parsed <- list()

  # 1. Parse + verify every file independently (a failing file does not stop the batch)
  for (i in seq_len(nrow(manifest))) {
    m <- manifest[i, ]
    cat("\n=== File", i, "of", nrow(manifest), ":", m$file_path, "\n")
    res <- tryCatch({
      doc <- read_statement_pdf(m$file_path, m$password)
      r <- resolve_bank(doc, m$bank_id, m$variant)
      if (is.null(r$bank)) stop_bsp("BANK_UNKNOWN", r$message)
      st <- r$bank$parse(doc)
      st <- complete_balances(st)
      ver <- verify_statement(st, tol)
      cat("Bank:", r$bank$id, "/", r$bank$variant, "- rows:", nrow(st$txns), "- status:", ver$status, "\n")
      list(ok = TRUE, bank = r$bank, st = st, ver = ver, message = r$message)
    }, error = function(e) {
      cat("ERROR:", conditionMessage(e), "\n")
      list(ok = FALSE, message = conditionMessage(e))
    })
    res$file_path <- m$file_path
    # per-file output folder (e.g. "same folder as the PDF") overrides the global one
    res$output_dir <- if (nzchar(m$output_dir)) m$output_dir else opts$output_dir
    parsed[[i]] <- res
  }

  consolidated <- opts$output_mode == "consolidated"

  # 2. Per-file outputs (or per-file verification only, when consolidating)
  for (res in parsed) {
    if (!res$ok) {
      results[[length(results) + 1]] <- result_row(res$file_path, message = res$message)
      next
    }
    out_file <- ""
    gid <- group_key(res$bank$id, res$st$meta, basename(res$file_path))
    if (!consolidated && !verify_only) {
      out_file <- output_path(res$output_dir, tools::file_path_sans_ext(basename(res$file_path)),
                              opts$on_exists)
      saved <- safe_write(res$st$txns, res$bank$id, out_file)
      if (!isTRUE(saved)) {
        results[[length(results) + 1]] <- result_row(res$file_path, res$bank, res$st, res$ver, "",
                                                     "ERROR", saved, gid)
        next
      }
    }
    results[[length(results) + 1]] <- result_row(res$file_path, res$bank, res$st, res$ver, out_file,
                                                 res$ver$status, res$message, gid)
    details[[length(details) + 1]] <- detail_rows(basename(res$file_path), out_file, res$st, res$ver)
  }

  # 3. Consolidated outputs: one per bank + account
  if (consolidated) {
    ok <- Filter(function(r) r$ok, parsed)
    keys <- vapply(ok, function(r) group_key(r$bank$id, r$st$meta, basename(r$file_path)), "")
    for (k in unique(keys)) {
      saved <- TRUE
      grp <- ok[keys == k]
      items <- lapply(grp, function(r) list(file_name = basename(r$file_path), st = r$st))
      cons <- consolidate_statements(items, tol)
      ver <- verify_statement(cons$st, tol)
      bank <- grp[[1]]$bank
      status <- if (ver$status == "ERROR" || any(cons$findings$severity == "ERROR")) "ERROR"
                else if (ver$status == "WARN" || any(cons$findings$severity == "WARN")) "WARN" else "OK"
      out_file <- ""
      if (!verify_only) {
        st <- cons$st
        base <- paste(bank$id, ifelse(is.na(st$meta$account_no), "NA", st$meta$account_no),
                      paste0(format(st$meta$period_start, "%Y%m%d"), "-", format(st$meta$period_end, "%Y%m%d")),
                      sep = "_")
        out_file <- output_path(grp[[1]]$output_dir, base, opts$on_exists)
        saved <- safe_write(st$txns, bank$id, out_file)
        if (!isTRUE(saved)) { status <- "ERROR"; out_file <- "" }
      }
      label <- paste0("[GABUNGAN ", length(cons$files), " file] ", paste(cons$files, collapse = " + "))
      results[[length(results) + 1]] <- result_row(label, bank, cons$st, ver, out_file, status,
                                                   if (isTRUE(saved)) "" else saved, k)
      details[[length(details) + 1]] <- detail_rows(label, out_file, cons$st, ver)
      if (nrow(cons$findings) > 0) {
        cons_details[[length(cons_details) + 1]] <- cbind(group_id = k, output_file = basename(out_file),
                                                          cons$findings)
      }
    }
  }

  empty_detail <- data.frame(file_name = character(), output_file = character(), excel_row = integer(),
                             date = character(), description = character(), debit = numeric(),
                             credit = numeric(), balance = numeric(), check = character(),
                             severity = character(), message = character())
  write_result_csv(do.call(rbind, results), file.path(work_dir, "run_result.csv"))
  d <- do.call(rbind, details)
  write_result_csv(if (is.null(d)) empty_detail else d, file.path(work_dir, "verification_detail.csv"))
  cd <- do.call(rbind, cons_details)
  write_result_csv(if (is.null(cd)) data.frame(group_id = character(), output_file = character(),
                                               file_name = character(), check = character(),
                                               severity = character(), message = character()) else cd,
                   file.path(work_dir, "consolidation_detail.csv"))
}

status <- tryCatch({ main(); "OK" }, error = function(e) {
  cat("FATAL:", conditionMessage(e), "\n")
  paste("FATAL:", conditionMessage(e))
})
finish(if (status == "OK") "OK" else "FATAL", status)
