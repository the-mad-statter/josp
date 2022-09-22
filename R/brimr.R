#' BRIMR Download Table 3
#'
#' @param year desired year
#' @param dest destination file
#'
#' @return destination file path
#' @export
#'
#' @description
#' The [Blue Ridge Institute for Medical Research](https://www.brimr.org/)
#' compiles and disseminates NIH funding data annually. This function attempts
#' to download Table 3 Medical Schools Only File 10 Mb; the comprehensive file
#' for all US Medical Schools receiving NIH funds for the desired year.
brimr_download_table_3 <-
  function(year, dest = sprintf("MedicalSchoolsOnly_%s.xls", year)) {
    file_name <- sprintf("MedicalSchoolsOnly_%s.xls", year)
    temp_file <- file.path(tempdir(), file_name)

    url <- sprintf(
      "http://www.brimr.org/NIH_Awards/%s/MedicalSchoolsOnly_%s.xls",
      year,
      year
    )

    r <- httr::GET(url, httr::write_disk(temp_file, TRUE))

    if (r$headers$`content-type` != "application/vnd.ms-excel")
      stop(sprintf("Data at %s does not appear to be an Excel document.", url))

    if (file.copy(temp_file, dest, TRUE))
      dest
  }

#' BRIMR Get Fig Data
#'
#' @param year desired year
#' @param dept desired department
#'
#' @return a tibble containing summary information regarding NIH funding
#' @export
#'
#' @examples
#' brimr_get_fig_data(2021, "PEDIATRICS")
brimr_get_fig_data <- function(year, dept) {
  y <- year

  raw <- josp::brimr_table_3 %>%
    dplyr::filter(year == y)

  awards_ranked_by_department <- raw %>%
    janitor::clean_names() %>%
    dplyr::filter(!is.na(pi_name)) %>%
    dplyr::group_by(organization_name, nih_dept_combining_name) %>%
    dplyr::summarize(
      total_award = sum(funding),
      n_grants = dplyr::n(),
      .groups = "drop"
    ) %>%
    dplyr::filter(nih_dept_combining_name == dept) %>%
    dplyr::arrange(dplyr::desc(total_award)) %>%
    tibble::rowid_to_column("rank")

  wusm_award_info <- awards_ranked_by_department %>%
    dplyr::filter(organization_name %in% c(
      "WASHINGTON UNIVERSITY",
      "WASHINGTON UNIVERSITY ST LOUIS"
    ))

  dplyr::tibble(
    year = y,
    fisc_year = sub("^\\d{2}", "FY", y),
    dept = dept,
    wusm_n_grants = wusm_award_info$n_grants,
    wusm_rank = wusm_award_info$rank,
    wusm_award = wusm_award_info$total_award,
    nih_total = sum(awards_ranked_by_department$total_award, na.rm = TRUE),
    nih_mean = mean(awards_ranked_by_department$total_award, na.rm = TRUE),
    non_wusm_award = nih_total - wusm_award,
    wusm_pct = wusm_award / nih_total
  )
}

#' BRIMR Make PowerPoint Slide
#'
#' @param year desired year
#' @param dept desired department
#' @param target destination file
#'
#' @export
brimr_make_slide <-
  function(year, dept, target = sprintf("brimr_%s_%s.pptx", dept, year)) {
  d <- dplyr::tibble(
    year = as.numeric(year):(as.numeric(year) - 3),
    dept = dept
  ) %>%
    purrr::pmap_dfr(brimr_get_fig_data)

  wusm_cagr <- scales::percent_format(0.1)(
    cagr(dplyr::last(d$wusm_award), dplyr::first(d$wusm_award), 3)
  )

  nih_cagr <- scales::percent_format(0.1)(
    cagr(dplyr::last(d$nih_total), dplyr::first(d$nih_total), 3)
  )

  d2_year_span <- sprintf(
    "%s-%s",
    sub("^\\d{2}", "", min(d$year)),
    sub("^\\d{2}", "", max(d$year))
  )

  df <- d %>%
    dplyr::select(year, wusm_rank, wusm_n_grants, wusm_pct) %>%
    dplyr::arrange(year) %>%
    dplyr::mutate(wusm_pct = scales::percent_format(0.1)(wusm_pct))

  ft <- dplyr::as_tibble(
    cbind(nms = names(df), t(df)),
    .name_repair = ~ c("nms", rev(d$fisc_year))
  ) %>%
    dplyr::mutate(dplyr::across(.fns = unlist)) %>%
    dplyr::filter(nms != "year") %>%
    dplyr::mutate(nms = c(
      "WUSM Dept. Rank",
      "WUSM Dept. # Grants",
      "WUSM % To Dept. NIH"
    )) %>%
    dplyr::rename(" " = nms) %>%
    flextable::flextable() %>%
    flextable::fontsize(size = 12) %>%
    flextable::width(1, 2) %>%
    flextable::width(2:5, 0.6)

  gg_d <- d %>%
    dplyr::mutate(
      gg_wusm_award = round(wusm_award / 1e+03, 0),
      gg_str_wusm_award = scales::dollar_format()(gg_wusm_award),
      gg_wusm_pct = 100 * 0.10 * wusm_pct * max(gg_wusm_award)
    )

  gg_d_gg_wusm_award_5k_ceiling <- ceiling(max(gg_d$gg_wusm_award) / 5000) * 5000
  gg_d_gg_wusm_award_break_labs <- seq(0, gg_d_gg_wusm_award_5k_ceiling, 5000)

  gg <- gg_d %>%
    ggplot2::ggplot(
      ggplot2::aes(year, gg_wusm_award, label = gg_str_wusm_award)
    ) +
    ggplot2::geom_col(color = "black", size = 1.5) +
    ggplot2::geom_point(ggplot2::aes(year, gg_wusm_pct), size = 3) +
    ggplot2::geom_line(ggplot2::aes(year, gg_wusm_pct), size = 1.5) +
    ggplot2::geom_text(nudge_y = 2000) +
    ggplot2::scale_x_continuous(
      name = ggplot2::element_blank(),
      breaks = d$year,
      labels = d$fisc_year
    ) +
    ggplot2::scale_y_continuous(
      name = "WUSM NIH Dept. Award $",
      limits = c(0, gg_d_gg_wusm_award_5k_ceiling),
      breaks = gg_d_gg_wusm_award_break_labs,
      labels = scales::dollar_format()(gg_d_gg_wusm_award_break_labs),
      sec.axis = ggplot2::sec_axis(
        ~ 100 * 0.10 * . / gg_d_gg_wusm_award_5k_ceiling,
        name = "WUSM % to NIH Dept. Total",
        breaks = seq(0, 10, 2),
        labels = scales::percent_format(0.1)(seq(0, 0.1, 0.02))
      )
    )

  pptx <- officer::read_pptx()

  pptx <- officer::add_slide(pptx, layout = "Two Content", master = "Office Theme")

  # title
  pptx <- officer::ph_with(
    x = pptx,
    value = "NIH WUSM Extramural Funding",
    location = officer::ph_location_type(type = "title")
  )

  # cagr
  fp_1 <- officer::fp_text(font.size = 14)
  bl_1 <- officer::block_list(officer::fpar(officer::ftext(
    sprintf("%s WUSM %s CAGR: %s", d2_year_span, dept, wusm_cagr),
    fp_1
  )))
  pptx <- officer::ph_with(
    x = pptx,
    value = bl_1,
    location = officer::ph_location(0.5, 0.5)
  )
  bl_2 <- officer::block_list(officer::fpar(officer::ftext(
    sprintf("%s NIH %s CAGR: %s", d2_year_span, dept, nih_cagr),
    fp_1
  )))
  pptx <- officer::ph_with(
    x = pptx,
    value = bl_2,
    location = officer::ph_location(0.5, 1)
  )

  # table
  pptx <- officer::ph_with(
    x = pptx,
    value = ft,
    location = officer::ph_location(0.5, 3.75)
  )

  # footnote
  fp_2 <- officer::fp_text(font.size = 8)
  bl_3 <- officer::block_list(officer::fpar(officer::ftext(paste0(
    "NIH Funding = federal fiscal years (October 1 = September 30); includes ",
    "both direct and indirect awards; includes research grants, training grants ",
    "and fellowships. Excludes R&D Contracts and ARRA funding."
  ), fp_2)))
  pptx <- officer::ph_with(
    x = pptx,
    value = bl_3,
    location = officer::ph_location(left = 0.5, top = 6.75, width = 9, height = 0.5)
  )
  bl_4 <- officer::block_list(officer::fpar(officer::ftext(
    "Source: Blue Ridge Institute for Medial Research website", fp_2
  )))
  pptx <- officer::ph_with(
    x = pptx,
    value = bl_4,
    location = officer::ph_location(left = 0.5, top = 7, width = 9, height = 0.5)
  )

  # ggplot
  pptx <- officer::ph_with(
    x = pptx,
    value = gg,
    location = officer::ph_location_right()
  )

  print(pptx, target = target)
}
