library(dplyr)

brimr_table_3 <- purrr::map_dfr(2011:2021, ~ {
  year <- .
  # if downloaded, they will need to be edited by hand
  # - award_notice_date -> yyyy-mm-dd
  #f <- file.path("data-raw", brimr_download_table_3(year))
  f <- file.path("data-raw", sprintf("MedicalSchoolsOnly_%s.xls", year))

  suppressWarnings(readxl::read_xls(path = f, sheet = 1, skip = 1)) %>%
    janitor::clean_names() %>%
    filter(!is.na(pi_name)) %>%
    mutate(year = year) %>%
    mutate_at(vars(one_of('zip_code')), as.character) %>%
    # edit files by hand to be custom yyyy-mm-dd
    mutate_at(vars(one_of('award_notice_date')), as.Date) %>%
    select(year, everything())
}) %>%
  mutate(funding = if_else(is.na(funding), award, funding)) %>%
  select(-award)

usethis::use_data(brimr_table_3, overwrite = TRUE)
