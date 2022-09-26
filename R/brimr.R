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

    if (r$headers$`content-type` != "application/vnd.ms-excel") {
      stop(sprintf("Data at %s does not appear to be an Excel document.", url))
    }

    if (file.copy(temp_file, dest, TRUE)) {
      dest
    }
  }

#' BRIMR Department Data
#'
#' @param year desired year
#' @param nih_dept_combining_name desired department
#'
#' @return a tibble containing summary information regarding NIH funding
#' @export
#'
#' @examples
#' brimr_dept_data(2021, "PEDIATRICS")
brimr_dept_data <- function(year, nih_dept_combining_name) {
  funding_ranked_by_department <- josp::brimr_table_3 %>%
    dplyr::filter(.data[["year"]] == .env[["year"]]) %>%
    dplyr::group_by(
      .data[["organization_name"]],
      .data[["nih_dept_combining_name"]]
    ) %>%
    dplyr::summarize(
      total_funding = sum(.data[["funding"]]),
      n_awards = dplyr::n(),
      .groups = "drop"
    ) %>%
    dplyr::filter(
      .data[["nih_dept_combining_name"]] ==
        .env[["nih_dept_combining_name"]]
    ) %>%
    dplyr::arrange(dplyr::desc(.data[["total_funding"]])) %>%
    tibble::rowid_to_column("rank")

  wusm_funding_info <- funding_ranked_by_department %>%
    dplyr::filter(.data[["organization_name"]] %in% c(
      "WASHINGTON UNIVERSITY",
      "WASHINGTON UNIVERSITY ST LOUIS"
    ))

  top_funding_info <- funding_ranked_by_department %>% dplyr::slice(1)

  dplyr::tibble(
    year = .env[["year"]],
    fisc_year = sub("^\\d{2}", "FY", .env[["year"]]),
    nih_dept_combining_name = .env[["nih_dept_combining_name"]],
    wusm_n_awards = wusm_funding_info$n_awards,
    wusm_rank = wusm_funding_info$rank,
    wusm_funding = wusm_funding_info$total_funding,
    nih_total_funding = sum(
      funding_ranked_by_department$total_funding,
      na.rm = TRUE
    ),
    nih_mean_funding = mean(
      funding_ranked_by_department$total_funding,
      na.rm = TRUE
    ),
    nih_non_wusm_funding = .data[["nih_total_funding"]] -
      .data[["wusm_funding"]],
    wusm_prop_nih_total_funding = .data[["wusm_funding"]] /
      .data[["nih_total_funding"]],
    top_organization_name = top_funding_info$organization_name,
    top_organization_funding = top_funding_info$total_funding
  )
}

#' BRIMR Map WUSM Department Name
#'
#' @param nih_dept_combining_name standardized NIH department name
#' @param sanitize replace spaces and ampersands
#'
#' @return the corresponding WUSM department name
#' @export
#'
#' @examples
#' brimr_map_wusm_dept("EMERGENCY MEDICINE")
brimr_map_wusm_dept <- function(nih_dept_combining_name, sanitize = FALSE) {
  wusm_dept <- josp::brimr_wusm_dept_mappings %>%
    dplyr::filter(
      .data[["nih_dept_combining_name"]] ==
        .env[["nih_dept_combining_name"]]
    ) %>%
    dplyr::pull(wusm_dept)

  if (sanitize) {
    wusm_dept <- gsub(" ", "_", wusm_dept)
    wusm_dept <- gsub("&", "and", wusm_dept)
  }

  return(wusm_dept)
}

#' BRIMR Slide Name
#'
#' @param nih_dept_combining_name department
#' @param year year
#'
#' @return a standardized slide name
#' @export
#'
#' @examples
#' brimr_dept_slide_name("PEDIATRICS", 2021)
brimr_dept_slide_name <- function(nih_dept_combining_name, year) {
  sprintf(
    "BRIMR_%s_%s.pptx",
    brimr_map_wusm_dept(nih_dept_combining_name, TRUE),
    year
  )
}

#' BRIMR Department Flextable
#'
#' @param data tibble containing years in rows and four variables at a minimum:
#' year, wusm_rank, wusm_n_awards, wusm_prop_nih_total_funding
#'
#' @return a flextable listing rank, number of grants, and percent of NIH
#' funding for each year
brimr_dept_flextable <- function(data) {
  data <- data %>%
    dplyr::select(
      .data[["year"]],
      .data[["wusm_rank"]],
      .data[["wusm_n_awards"]],
      .data[["wusm_prop_nih_total_funding"]]
    ) %>%
    dplyr::arrange(.data[["year"]]) %>%
    dplyr::mutate(
      wusm_prop_nih_total_funding = scales::percent_format(0.1)(
        .data[["wusm_prop_nih_total_funding"]]
      )
    )

  dplyr::as_tibble(
    cbind(nms = names(data), t(data)),
    .name_repair = ~ c("nms", paste0("FY", sub("^\\d{2}", "", data$year)))
  ) %>%
    dplyr::mutate(dplyr::across(.fns = unlist)) %>%
    dplyr::filter(.data[["nms"]] != "year") %>%
    dplyr::mutate(nms = c(
      "WUSM Dept. Rank",
      "WUSM Dept. # Grants",
      "WUSM % To Dept. NIH"
    )) %>%
    dplyr::rename(" " = .data[["nms"]]) %>%
    flextable::flextable() %>%
    flextable::fontsize(size = 12) %>%
    flextable::width(1, 2) %>%
    flextable::width(2:5, 0.6)
}

#' BRIMR Department Slide
#'
#' @param year desired year
#' @param nih_dept_combining_name desired department
#' @param target path to the pptx file to write
#'
#' @export
brimr_dept_slide <-
  function(year,
           nih_dept_combining_name,
           target = brimr_dept_slide_name(nih_dept_combining_name, year)) {
    d <- dplyr::tibble(
      year = as.numeric(year):(as.numeric(year) - 3),
      nih_dept_combining_name = nih_dept_combining_name
    ) %>%
      purrr::pmap_dfr(brimr_dept_data)

    wusm_cagr <- scales::percent_format(0.1)(
      cagr(
        dplyr::last(d$wusm_funding), dplyr::first(d$wusm_funding), 3
      )
    )

    nih_cagr <- scales::percent_format(0.1)(
      cagr(
        dplyr::last(d$nih_total_funding), dplyr::first(d$nih_total_funding), 3
      )
    )

    year_span <- sprintf(
      "%s-%s",
      sub("^\\d{2}", "", min(d$year)),
      sub("^\\d{2}", "", max(d$year))
    )

    ft <- brimr_dept_flextable(d)

    gg_d <- d %>%
      dplyr::mutate(
        gg_wusm_funding = round(.data[["wusm_funding"]] / 1e+03, 0),
        gg_wusm_funding_str = scales::dollar_format()(
          .data[["gg_wusm_funding"]]
        ),
        gg_wusm_prop_nih_total_funding =
          100 * .data[["wusm_prop_nih_total_funding"]] *
            0.10 * max(.data[["gg_wusm_funding"]])
      )

    gg <- gg_d %>%
      ggplot2::ggplot(
        ggplot2::aes(
          .data[["year"]],
          .data[["gg_wusm_funding"]],
          label = .data[["gg_wusm_funding_str"]]
        )
      ) +
      ggplot2::geom_col(color = "black", size = 1.5) +
      ggplot2::geom_point(
        ggplot2::aes(year, .data[["gg_wusm_prop_nih_total_funding"]]),
        size = 3
      ) +
      ggplot2::geom_line(
        ggplot2::aes(year, .data[["gg_wusm_prop_nih_total_funding"]]),
        size = 1.5
      ) +
      ggplot2::geom_text(nudge_y = 2000) +
      ggplot2::scale_x_continuous(
        name = ggplot2::element_blank(),
        breaks = d$year,
        labels = d$fisc_year
      )

    y_max <- ggplot2::ggplot_build(gg)$layout$panel_params[[1]]$y.range[2]

    gg <- gg +
      ggplot2::scale_y_continuous(
        name = "WUSM NIH Dept. Award $",
        limits = c(0, y_max),
        labels = scales::label_dollar(),
        sec.axis = ggplot2::sec_axis(
          ~ 100 * 0.10 * . / y_max,
          name = "WUSM % to NIH Dept. Total",
          breaks = seq(0, 10, 2),
          labels = scales::percent_format(0.1)(seq(0, 0.1, 0.02))
        )
      )

    pptx <- officer::read_pptx()

    pptx <- officer::add_slide(
      pptx,
      layout = "Two Content",
      master = "Office Theme"
    )

    # title
    pptx <- officer::ph_with(
      x = pptx,
      value = "NIH WUSM Extramural Funding",
      location = officer::ph_location_type(type = "title")
    )

    # cagr
    fp_1 <- officer::fp_text(font.size = 14)
    bl_1 <- officer::block_list(officer::fpar(officer::ftext(
      sprintf(
        "%s WUSM %s CAGR: %s",
        year_span,
        brimr_map_wusm_dept(nih_dept_combining_name),
        wusm_cagr
      ),
      fp_1
    )))
    pptx <- officer::ph_with(
      x = pptx,
      value = bl_1,
      location = officer::ph_location(0.5, 0.5)
    )
    bl_2 <- officer::block_list(officer::fpar(officer::ftext(
      sprintf(
        "%s NIH %s CAGR: %s",
        year_span,
        brimr_map_wusm_dept(nih_dept_combining_name),
        nih_cagr
      ),
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
      "NIH Funding = federal fiscal years (October 1 = September 30); ",
      "includes both direct and indirect awards; includes research grants, ",
      "training grants and fellowships. Excludes R&D Contracts and ARRA ",
      "funding."
    ), fp_2)))
    pptx <- officer::ph_with(
      x = pptx,
      value = bl_3,
      location = officer::ph_location(
        left = 0.5,
        top = 6.75,
        width = 9,
        height = 0.5
      )
    )
    bl_4 <- officer::block_list(officer::fpar(officer::ftext(
      "Source: Blue Ridge Institute for Medial Research website", fp_2
    )))
    pptx <- officer::ph_with(
      x = pptx,
      value = bl_4,
      location = officer::ph_location(
        left = 0.5,
        top = 7,
        width = 9,
        height = 0.5
      )
    )

    # ggplot
    pptx <- officer::ph_with(
      x = pptx,
      value = gg,
      location = officer::ph_location_right()
    )

    print(pptx, target = target)
  }

#' BRIMR Ranking Table Row
#'
#' @param year desired year
#' @param nih_dept_combining_name the NIH combining name of the department
#'
#' @return a tibble containing data for one row in a ranking table
#'
#' @examples
#' josp:::brimr_ranking_table_row(2021, "PEDIATRICS")
brimr_ranking_table_row <- function(year, nih_dept_combining_name) {
  year_1 <- brimr_dept_data(year, nih_dept_combining_name)
  year_0 <- brimr_dept_data(year - 1, nih_dept_combining_name)

  wusm_dept <- josp::brimr_wusm_dept_mappings %>%
    dplyr::filter(
      .data[["nih_dept_combining_name"]] == .env[["nih_dept_combining_name"]]
    ) %>%
    dplyr::pull(.data[["wusm_dept"]])

  tibble::tibble(
    wusm_dept = wusm_dept,
    nih_dept_combining_name = nih_dept_combining_name,
    wusm_rank_status = dplyr::case_when(
      year_0$wusm_rank < year_1$wusm_rank ~ "down", # down in ranking
      year_0$wusm_rank == year_1$wusm_rank ~ "equal",
      year_0$wusm_rank > year_1$wusm_rank ~ "up", # up in ranking
      TRUE ~ NA_character_
    ),
    wusm_rank_1 = year_1$wusm_rank,
    wusm_rank_0 = year_0$wusm_rank,
    wusm_funding_1 = year_1$wusm_funding,
    wusm_funding_0 = year_0$wusm_funding,
    top_organization_name_1 = year_1$top_organization_name,
    top_organization_funding_1 = year_1$top_organization_funding,
    top_organization_name_0 = year_0$top_organization_name,
    top_organization_funding_0 = year_0$top_organization_funding
  )
}

#' BRIMR Ranking Table
#'
#' @param year desired year to be compared to the previous
#'
#' @return a tibble containing data that compares WUSM departments ranks and
#' funding vs the top funded department for the provided and previous year.
#' @export
#'
#' @examples
#' brimr_ranking_table(2021)
brimr_ranking_table <- function(year) {
  purrr::pmap_dfr(
    tibble::tibble(
      year = year,
      nih_dept_combining_name =
        josp::brimr_wusm_dept_mappings$nih_dept_combining_name
    ),
    brimr_ranking_table_row
  )
}

#' BRIMR Ranking Flextable
#'
#' @param year desired year to be compared to the previous
#'
#' @return a printable flextable version of brimr_ranking_table(year)
#' @export
#'
#' @examples
#' brimr_ranking_flextable(2021)
brimr_ranking_flextable <- function(year) {
  year_0 <- year - 1
  year_1 <- year

  rt <- brimr_ranking_table(year) %>%
    dplyr::select(-.data[["nih_dept_combining_name"]]) %>%
    dplyr::arrange(.data[["wusm_dept"]])

  k <- list(
    "light_red"    = "#ffcccc",
    "yellow" = "#fff2cc",
    "light_green"  = "#e2efda",
    "light_blue"   = "#d9e1f2",
    "dark_green" = "#008000",
    "dark_red" = "#ff0000",
    "dark_blue" = "#0000ff"
  )

  rank_status_rows <- list(
    "down" = which(rt$wusm_rank_status == "down"),
    "equal" = which(rt$wusm_rank_status == "equal"),
    "up" = which(rt$wusm_rank_status == "up")
  )

  rank_2_rows <- which(rt$wusm_rank_1 == 2)

  rt %>%
    dplyr::mutate(
      dplyr::across(
        c(
          dplyr::matches("wusm_funding_[01]"),
          dplyr::matches("top_organization_funding_[01]")
        ),
        scales::label_dollar()
      ),
      wusm_rank_status = dplyr::case_when(
        wusm_rank_status == "up" ~ "\u2191",
        wusm_rank_status == "equal" ~ "=",
        wusm_rank_status == "down" ~ "\u2193",
        TRUE ~ NA_character_
      )
    ) %>%
    flextable::flextable() %>%
    flextable::bg(i = rank_status_rows$down, bg = k$light_red) %>%
    flextable::bg(i = rank_status_rows$equal, bg = k$yellow) %>%
    flextable::bg(i = rank_status_rows$up, bg = k$light_green) %>%
    flextable::bg(i = which(rt$wusm_rank_1 == 1), bg = k$light_blue) %>%
    flextable::color(i = rank_status_rows$down, j = 2, color = k$dark_red) %>%
    flextable::color(i = rank_status_rows$up, j = 2, color = k$dark_green) %>%
    flextable::color(i = rank_2_rows, j = 3, color = k$dark_blue) %>%
    flextable::set_header_labels(
      wusm_dept = "",
      wusm_rank_status = paste(
        sub("^\\d{2}", "", year_1),
        "vs",
        sub("^\\d{2}", "", year_0)
      ),
      wusm_rank_1 = year_1,
      wusm_rank_0 = year_0,
      wusm_funding_1 = year_1,
      wusm_funding_0 = year_0,
      top_organization_name_1 = "#1 in Funding",
      top_organization_funding_1 = "Amount",
      top_organization_name_0 = "#1 in Funding",
      top_organization_funding_0 = "Amount"
    ) %>%
    flextable::add_header_row(
      values = c("Department", "Rank", "Funding", year_1, year_0),
      colwidths = c(1, 3, 2, 2, 2)
    ) %>%
    flextable::add_header_row(
      values = c("WUSM", "NIH"),
      colwidths = c(6, 4)
    ) %>%
    flextable::align(align = "center", part = "all") %>%
    flextable::align(j = 1, align = "left", part = "body") %>%
    flextable::fontsize(size = 9, part = "all") %>%
    flextable::padding(
      padding.top = 5,
      padding.bottom = 5,
      padding.left = 0,
      padding.right = 0,
      part = "all"
    ) %>%
    flextable::width(j = 1,              width = 1.20) %>% # wusm depts
    flextable::width(j = 2,              width = 0.60) %>% # wusm rank change
    flextable::width(j = c(3, 4),        width = 0.50) %>% # wusm ranks
    flextable::width(j = c(5, 6, 8, 10), width = 1.00) %>% # amounts
    flextable::width(j = c(7, 9),        width = 3.00)     # org names
}

#' BRIMR Ranking Slide
#'
#' @param year desired year to be compared to the previous
#' @param target path to the pptx file to write
#'
#' @export
brimr_ranking_slide <-
  function(year, target = sprintf("BRIMR_Department_Rankings_%s.pptx", year)) {
    template_16x9 <- system.file("officer/template_16x9.pptx", package = "josp")
    pptx <- officer::read_pptx(template_16x9)

    pptx <- officer::add_slide(
      pptx,
      layout = "Blank",
      master = "Office Theme"
    )

    pptx <- officer::ph_with(
      x = pptx,
      value = brimr_ranking_flextable(year),
      location = officer::ph_location(
        left = 0.25,
        top = 0.25,
        width = 15.5,
        height = 8.5
      )
    )

    print(pptx, target = target)
  }
