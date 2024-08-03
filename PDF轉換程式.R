#!/usr/bin/env RScript

pacman::p_load("pdftools")
pacman::p_load("tidyverse")
pacman::p_load("openxlsx")

parser <- function(file_name) {

  find_subjects <- function(r) {
    str_match_all(r, r"((.{1,2}x)[0-9\\.]{1,4})")[[1]][, 2] |>
      trimws() |>
      str_replace("x$", "")
  }

  row_parser <- function(subjects) {
    pp <- paste0(subjects, collapse="|")
    pw <- paste0("[", pp, "]x[0-9\\.]{1,4}")
    pc <- paste0("[", pp, "] [1-9][0-9]")

    sub_list <- as.list(rep(0,  length(subjects)))
    names(sub_list) <- subjects

    cmp_list <- as.list(rep("", length(subjects)))
    names(cmp_list) <- paste("同分", subjects, sep = "_")

    f <- function(rec) {
      v <- str_split(rec, " +")[[1]]
      # 採計及加權
      vm <- str_detect(v, pw)
      sl <- sub_list
      for (w in v[vm]) {
        p <- str_split(w, "x")[[1]]
        sl[p[1]] <- as.numeric(p[2])
      }
      # 同分
      cl <- cmp_list
      for (cc in v[str_detect(v, pc)]) {
        p <- str_split(cc, " ")[[1]]
        cl[paste0("同分_", p[1])] <- p[2]
      }

      list(
        data.frame(
          編號 = v[1],
          學校 = v[2],
          科系 = v[3],
          錄取人數 = v[max(which(vm)) + 1],
          錄取分數 = v[max(which(vm)) + 2],
          原住民   = v[length(v) - 4],
          退伍軍人 = v[length(v) - 3],
          僑生     = v[length(v) - 2],
          蒙藏生   = v[length(v) - 1],
          派外子女 = v[length(v)]
        ),
        data.frame(sl),
        data.frame(cl)
      ) |>
        reduce(cbind)
    }
    return(f)
  }

  txt <- pdf_text(file_name)    # per page
  subjects <- txt |> map(find_subjects) |> unlist() |> unique()
  rp <- row_parser(subjects)

  dat <- txt |>
    imap(function(x, y) {
      d <- x |> str_split("\n") |> unlist()
      data.frame(頁號 = y, 虛擬列號 = seq_along(d), record = d) |>
        filter(record != "") |>
        filter(str_detect(record, r"(第\d頁)", TRUE)) |>
        filter(str_detect(record, r"(系組)", TRUE) & str_detect(record, r"(^\d+ )")) |>
        mutate(列號 = row_number())
    }) |>
    reduce(rbind)
  out <- dat$record |>
    map(rp) |>
    reduce(rbind) |>
    cbind(dat) |>
    rename(原始資料 = record)

  invisible(out)
}

dat112 <- parser("data/112_result_school_data.pdf")

dat112 |>
  select(-原始資料, -虛擬列號) |>
  write.xlsx("output/112.xlsx", asTable=TRUE)
