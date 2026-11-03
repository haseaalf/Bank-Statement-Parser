# BSP engine tests.  Usage (from the repository folder):
#   "C:\Program Files\R\R-4.2.2\bin\Rscript.exe" tests\run_tests.R
#
# 1. Golden test  : BRI sample PDF must reproduce "input output sample/*.xlsx"
# 2. Sample banks : every PDF in the sample folders must be detected as the
#                   expected bank and pass verification without ERROR
# Sample statements are confidential and not in git; missing folders are skipped.

repo <- normalizePath(file.path(dirname(sub("--file=", "", grep("--file=", commandArgs(), value = TRUE))), ".."))
R_script_path <- file.path(repo, "BSP_modular_engine")
for (folder in c("Core", "Banks")) {
  for (f in sort(list.files(file.path(R_script_path, folder), pattern = "\\.R$", full.names = TRUE))) source(f)
}
out_dir <- file.path(repo, "tests", "output")
dir.create(out_dir, showWarnings = FALSE)

failures <- 0
check <- function(ok, label) {
  cat(if (ok) "  PASS " else "  FAIL ", label, "\n", sep = "")
  if (!ok) failures <<- failures + 1
}

run_file <- function(pdf, bank_id = "AUTO") {
  doc <- read_statement_pdf(pdf)
  r <- resolve_bank(doc, bank_id)
  if (is.null(r$bank)) return(list(bank = NULL, resolve = r))
  st <- complete_balances(r$bank$parse(doc))
  list(bank = r$bank, resolve = r, st = st, ver = verify_statement(st))
}

# ---- 1. Golden test -----------------------------------------------------------
cat("\n[1] Golden test BRI\n")
sample_dir <- file.path(repo, "input output sample")
pdfs <- list.files(sample_dir, pattern = "\\.pdf$", full.names = TRUE)
if (length(pdfs) == 0) {
  cat("  SKIP (folder 'input output sample' tidak ada)\n")
} else {
  for (pdf in pdfs) {
    expected_file <- sub("\\.pdf$", ".xlsx", pdf)
    res <- run_file(pdf)
    out_file <- file.path(out_dir, basename(expected_file))
    write_output_xlsx(res$st$txns, res$bank$id, out_file)

    exp <- openxlsx::read.xlsx(expected_file, colNames = FALSE)
    got <- openxlsx::read.xlsx(out_file, colNames = FALSE)
    check(identical(dim(exp), dim(got)), paste("dimensi sama", paste(dim(got), collapse = "x")))
    check(identical(unlist(exp[1, ]), unlist(got[1, ])), "header identik")
    check(identical(openxlsx::getSheetNames(expected_file), openxlsx::getSheetNames(out_file)), "nama sheet identik")
    e <- exp[-1, ]; g <- got[-1, ]
    check(identical(e[[1]], g[[1]]), "kolom tanggal identik")
    for (k in 3:5) check(isTRUE(all.equal(as.numeric(e[[k]]), as.numeric(g[[k]]))),
                         paste("kolom", c("", "", "Debit", "Credit", "Balance")[k], "identik"))
    norm <- function(x) gsub("\\s+", "", x)
    diff_rows <- which(e[[2]] != g[[2]])
    diff_norm <- which(norm(e[[2]]) != norm(g[[2]]))
    cat("  INFO deskripsi identik persis:", nrow(e) - length(diff_rows), "dari", nrow(e),
        "| beda setelah abaikan spasi:", length(diff_norm), "\n")
    for (i in head(diff_norm, 10)) {
      cat("       baris", i + 1, ": sampel =", shQuote(gsub("\n", "\\\\n", e[[2]][i])),
          "| hasil =", shQuote(gsub("\n", "\\\\n", g[[2]][i])), "\n")
    }
    check(res$ver$status == "OK", paste("verifikasi", res$ver$status))
    # Excel refuses packages with folder entries or parts that point to missing files
    zl <- zip::zip_list(out_file)$filename
    rels <- paste(readLines(unz(out_file, "xl/worksheets/_rels/sheet1.xml.rels"), warn = FALSE), collapse = "")
    check(!any(grepl("/$", zl)) && zl[1] == "[Content_Types].xml" && !grepl("drawing", rels),
          "paket xlsx bersih (tanpa entri folder / relasi ke file yang tidak ada)")
  }
}

# ---- 2. All sample banks ------------------------------------------------------
cat("\n[2] Sample bank statements\n")
expected <- utils::read.csv(file.path(repo, "tests", "expected_banks.csv"), stringsAsFactors = FALSE)
all_pdfs <- c(pdfs, list.files(file.path(repo, "other bank statement sample"), pattern = "\\.pdf$",
                               full.names = TRUE))
only <- commandArgs(trailingOnly = TRUE)   # optional file name filter, e.g. "Mandiri"
if (length(only) > 0) all_pdfs <- all_pdfs[grepl(only[1], basename(all_pdfs), ignore.case = TRUE)]
for (pdf in all_pdfs) {
  name <- basename(pdf)
  exp_id <- expected$bank_id[vapply(expected$pattern, function(p) grepl(p, name, ignore.case = TRUE), TRUE)]
  cat(" -", name, "\n")
  res <- tryCatch(run_file(pdf), error = function(e) { cat("    error:", conditionMessage(e), "\n"); NULL })
  if (is.null(res)) { check(FALSE, "parse"); next }
  got_id <- if (is.null(res$bank)) "(tidak dikenali)" else res$bank$id
  check(length(exp_id) == 1 && identical(got_id, exp_id[1]),
        paste0("deteksi = ", got_id, " (harapan ", paste(exp_id, collapse = "/"), ", skor ",
               round(res$resolve$score, 2), ")"))
  sc <- score_banks(read_statement_pdf(pdf))
  check(nrow(sc) < 2 || sc$score[2] < DETECT_MIN_SCORE,
        paste0("tidak ambigu (skor kedua ", sc$bank_id[2], " = ", round(sc$score[2], 2), ")"))
  if (is.null(res$bank)) next
  write_output_xlsx(res$st$txns, res$bank$id, file.path(out_dir, sub("\\.pdf$", ".xlsx", name)))
  v <- res$ver
  check(v$status != "ERROR", sprintf("verifikasi %s: %d baris, awal %s (%s), akhir %s, debit %s, kredit %s",
                                     v$status, v$summary$rows, fmt_num(v$summary$opening), v$summary$opening_source,
                                     fmt_num(v$summary$closing), fmt_num(v$summary$total_debit),
                                     fmt_num(v$summary$total_credit)))
  bad <- v$findings[v$findings$severity %in% c("ERROR", "WARN"), ]
  for (i in seq_len(min(8, nrow(bad)))) cat("       ", bad$severity[i], "baris", bad$row[i], ":", bad$message[i], "\n")
}

# ---- 3. Consolidation logic (statements built by splitting a parsed sample) ----
cat("\n[3] Konsolidasi\n")
bca <- list.files(file.path(repo, "other bank statement sample"), pattern = "^BCA.*\\.pdf$", full.names = TRUE)
if (length(bca) == 0) {
  cat("  SKIP (sampel BCA tidak ada)\n")
} else {
  full <- run_file(bca[1])$st
  t <- full$txns
  split_st <- function(rows, ps, pe) {
    tt <- t[rows, ]
    open <- if (rows[1] == 1) full$opening$value else t$balance[rows[1] - 1]
    st <- new_statement(tt, meta = modifyList(full$meta, list(period_start = ps, period_end = pe)),
                        stated = list(opening = open, closing = tail(tt$balance, 1)))
    st$txns$balance_source <- tt$balance_source
    complete_balances(st)
  }
  cut1 <- max(which(t$date <= as.Date("2026-02-10")))
  cut2 <- max(which(t$date <= as.Date("2026-02-20")))
  p1 <- split_st(1:cut1, as.Date("2026-02-01"), as.Date("2026-02-10"))
  p2 <- split_st((cut1 + 1):cut2, as.Date("2026-02-11"), as.Date("2026-02-20"))
  p3 <- split_st((cut2 + 1):nrow(t), as.Date("2026-02-21"), as.Date("2026-02-28"))

  # given in random order -> must be sorted, continuous, and equal to the full statement
  c1 <- consolidate_statements(list(list(file_name = "c.pdf", st = p3), list(file_name = "a.pdf", st = p1),
                                    list(file_name = "b.pdf", st = p2)))
  v1 <- verify_statement(c1$st)
  check(nrow(c1$findings) == 0 && v1$status == "OK", "3 periode digabung: urut, saldo bersambung, verifikasi OK")
  check(isTRUE(all.equal(c1$st$txns$balance, t$balance)) && identical(c1$files, c("a.pdf", "b.pdf", "c.pdf")),
        "hasil gabungan = statement asli")

  c2 <- consolidate_statements(list(list(file_name = "a.pdf", st = p1), list(file_name = "c.pdf", st = p3)))
  check(any(c2$findings$check == "gap") && any(c2$findings$check == "continuity" & c2$findings$severity == "ERROR"),
        "periode hilang terdeteksi (gap + saldo tidak bersambung)")

  c3 <- consolidate_statements(list(list(file_name = "a.pdf", st = p1), list(file_name = "a copy.pdf", st = p1),
                                    list(file_name = "b.pdf", st = p2)))
  check(any(c3$findings$check == "duplicate") && identical(c3$files, c("a.pdf", "b.pdf")),
        "file duplikat terdeteksi dan tidak ikut digabung")
}

# ---- 4. Engine end-to-end (as called by the Excel macro) ---------------------
cat("\n[4] Engine end-to-end\n")
if (length(all_pdfs) == 0 || length(only) > 0) {
  cat("  SKIP\n")
} else {
  rscript <- file.path(R.home("bin"), "Rscript.exe")
  work <- file.path(out_dir, "work"); dir.create(work, showWarnings = FALSE)
  outx <- file.path(work, "out"); unlink(outx, recursive = TRUE); dir.create(outx)
  run_engine <- function(mode, files, opts) {
    utils::write.csv(data.frame(file_path = files, bank_id = "AUTO", variant = "", password = ""),
                     file.path(work, "job_manifest.csv"), row.names = FALSE, fileEncoding = "UTF-8")
    utils::write.csv(data.frame(key = names(opts), value = unlist(opts)),
                     file.path(work, "job_options.csv"), row.names = FALSE, fileEncoding = "UTF-8")
    system2(rscript, shQuote(c(file.path(R_script_path, "BSP_engine.R"), work, R_script_path, mode)),
            stdout = FALSE, stderr = FALSE)
    readLines(file.path(work, "engine_done.txt"))[1]
  }
  check(run_engine("list_banks", all_pdfs[1], list()) == "OK" &&
          nrow(read.csv(file.path(work, "bank_registry.csv"))) >= 9, "list_banks menulis bank_registry.csv")
  check(run_engine("detect", all_pdfs, list()) == "OK", "detect selesai")
  d <- read.csv(file.path(work, "detect_result.csv"), stringsAsFactors = FALSE)
  check(nrow(d) == length(all_pdfs) && all(d$bank_id != ""), "semua file terdeteksi di detect_result.csv")
  check(run_engine("parse", all_pdfs, list(output_dir = outx, output_mode = "per_file", on_exists = "suffix")) == "OK",
        "parse per-file selesai")
  r <- read.csv(file.path(work, "run_result.csv"), stringsAsFactors = FALSE)
  check(all(r$status %in% c("OK", "WARN")) && all(file.exists(r$output_file)),
        paste("semua output per-file dibuat:", length(list.files(outx)), "file"))
  fake <- file.path(work, "rusak.pdf"); writeLines("bukan pdf", fake)
  check(run_engine("parse", c(fake, all_pdfs[1]), list(output_dir = outx, output_mode = "per_file",
                                                     on_exists = "overwrite")) == "OK", "batch dengan file rusak tetap selesai")
  r <- read.csv(file.path(work, "run_result.csv"), stringsAsFactors = FALSE)
  check(r$status[1] == "ERROR" && grepl("tidak bisa dibaca", r$message[1]) && r$status[2] %in% c("OK", "WARN"),
        paste0("file rusak ditandai ERROR (", r$message[1], "), file lain tetap diproses"))
  dup <- c(all_pdfs[1], all_pdfs[1])
  check(run_engine("parse", dup, list(output_dir = outx, output_mode = "consolidated", on_exists = "suffix")) == "OK",
        "parse konsolidasi selesai")
  r <- read.csv(file.path(work, "run_result.csv"), stringsAsFactors = FALSE)
  cd <- read.csv(file.path(work, "consolidation_detail.csv"), stringsAsFactors = FALSE)
  check(sum(grepl("^\\[GABUNGAN", r$file_name)) == 1 && any(cd$check == "duplicate"),
        "konsolidasi: 1 output per rekening, duplikat ditandai")

  # output names: per-file original / per-file standard / consolidated is always standard
  bri <- all_pdfs[grepl("MUNTILAN", basename(all_pdfs))][1]
  if (!is.na(bri)) {
    std <- "BRI_025101000147560_20240801-20240831.xlsx"
    run_engine("parse", bri, list(output_dir = outx, output_mode = "per_file", on_exists = "overwrite",
                                  name_style = "standard"))
    r <- read.csv(file.path(work, "run_result.csv"), stringsAsFactors = FALSE)
    check(basename(r$output_file[1]) == std, paste("nama standar per file:", basename(r$output_file[1])))
    run_engine("parse", bri, list(output_dir = outx, output_mode = "per_file", on_exists = "overwrite",
                                  name_style = "original"))
    r <- read.csv(file.path(work, "run_result.csv"), stringsAsFactors = FALSE)
    check(basename(r$output_file[1]) == sub("\\.pdf$", ".xlsx", basename(bri)), "nama asli PDF per file")
    run_engine("parse", bri, list(output_dir = outx, output_mode = "consolidated", on_exists = "overwrite",
                                  name_style = "original"))
    r <- read.csv(file.path(work, "run_result.csv"), stringsAsFactors = FALSE)
    g <- r[grepl("^\\[GABUNGAN", r$file_name), ]
    check(nrow(g) == 1 && basename(g$output_file) == std, "gabungan selalu memakai nama standar")
  }
}

# ---- 5. Long names (regression: "FATAL: path too long" with 6 consolidated files) ----
cat("\n[5] Nama file panjang\n")
long <- paste0("C:\\folder\\", strrep("!!!!!!! MUNTILAN ARIE 2025 1_025101000147560_2024-08-01_2024-08-31_00290827 ", 8), ".pdf")
row <- tryCatch(result_row(long), error = function(e) NULL)
check(!is.null(row) && nchar(row$file_name) > 260, paste("result_row menerima path", nchar(long), "karakter"))
names6 <- paste0(strrep("x", 90), 1:6, ".pdf")
row <- tryCatch(result_row("", file_name = "[GABUNGAN 6 file] a s/d f", source_files = paste(names6, collapse = " | ")),
                error = function(e) NULL)
check(!is.null(row) && length(strsplit(row$source_files, " \\| ")[[1]]) == 6, "baris gabungan 6 file dengan source_files")

# ---- 6. Monthly recap (rekap mutasi) ------------------------------------------------
cat("\n[6] Rekap mutasi\n")
ita <- file.path(repo, "mutasi_ITA YUNI ASTUTI", "668001019813539")
if (!dir.exists(ita) || length(all_pdfs) == 0 || length(only) > 0) {
  cat("  SKIP (folder mutasi_ITA YUNI ASTUTI tidak ada)\n")
} else {
  ita_pdfs <- sort(list.files(ita, pattern = "[.]pdf$", full.names = TRUE))
  # skip June 2026 (missing month) and upload March 2026 twice (duplicate)
  use <- c(ita_pdfs[!grepl("20260601", ita_pdfs)], ita_pdfs[grepl("20260301", ita_pdfs)], all_pdfs)
  parsed <- lapply(use, function(p) {
    r <- run_file(p)
    list(ok = TRUE, bank = r$bank, st = r$st, file_path = p)
  })
  # the second March upload gets its own name, as a re-uploaded copy would
  dup_i <- length(ita_pdfs)
  parsed[[dup_i]]$file_path <- sub("[.]pdf$", " (copy).pdf", parsed[[dup_i]]$file_path)
  rc <- build_recap(parsed)
  acc <- Filter(function(a) identical(a$account_no, "668001019813539"), rc$accounts)[[1]]
  tb <- acc$table
  w <- function(m) tb$warning[tb$month == m]
  names_n <- length(unique(vapply(rc$accounts, function(a) a$account_name %||% "", "")))
  check(length(rc$accounts) == length(all_pdfs) + 1 && names_n >= 3,
        paste0("banyak bank/rekening/nasabah: ", length(rc$accounts), " rekening, ", names_n, " nasabah, ",
               length(rc$months), " bulan"))
  check(!tb$has_data[tb$month == "2026-06"] && grepl("tidak ada", w("2026-06")), "bulan hilang (Jun 2026) ditandai")
  check(grepl("tidak bersambung", w("2026-07")), "bulan sesudahnya (Jul 2026) diberi warning urutan")
  check(grepl("Duplikat", w("2026-03")), "file duplikat (Mar 2026) diberi warning")
  check(!any(grepl("Saldo bulan lalu", tb$warning)), "saldo bulan lalu + kredit - debit = saldo di semua bulan")
  check(isTRUE(all.equal(tb$saldo[tb$month == "2026-10"], 27516964)), "saldo akhir Okt 2026 = 27.516.964")
  one <- Filter(function(a) a$bank_id == "BCA", rc$accounts)[[1]]$table
  check(sum(one$has_data) == 1 && !any(nzchar(one$warning)),
        "rekening 1 bulan: bulan di luar periodenya kosong tanpa warning")
  f <- file.path(out_dir, "rekap_test.xlsx")
  write_recap_xlsx(rc, f)
  sh <- openxlsx::getSheetNames(f)
  check(sh[1] == "REKAP" && length(sh) == length(rc$accounts) + 1 && "668001019813539" %in% sh,
        "workbook rekap: REKAP + 1 sheet per rekening")
  x <- openxlsx::read.xlsx(f, sheet = "668001019813539", colNames = FALSE, skipEmptyCols = FALSE)
  check(ncol(x) >= 5 && any(grepl("PERINGATAN", x[[5]])), "warning tertulis di kolom E sheet rekening")
}

cat("\n", if (failures == 0) "SEMUA TES LULUS" else paste(failures, "TES GAGAL"), "\n")
quit(status = if (failures == 0) 0 else 1)
