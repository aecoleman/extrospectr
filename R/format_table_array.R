#' Format Table Array
#'
#' @param tbl_ary list object
#'
#' @return data.frame
#'
#' @importFrom purrr map map_if transpose map_dfr
#'
.format_table_array <- function(tbl_ary) {

  tbl_ary %>%
    # Replace any NULL elements with NA, since their NULLity will cause issues
    # when we try to turn the information into a data.frame
    purrr::map(
      ~ .x %>%
        purrr::map_if(
          .p = is.null,
          .f = ~ NA)
      ) %>%
    purrr::transpose() %>%
    purrr::map_dfr(
      ~ data.frame(
          EntryID              = .x[[1]],
          Subject              = .x[[2]],
          CreationTime         = .x[[3]] %>% .COMDate_to_POSIX(),
          LastModificationTime = .x[[4]] %>% .COMDate_to_POSIX(),
          MessageClass         = .x[[5]],
          Sender               = .x[[6]],
          ReceivedTime         = .x[[7]] %>% .COMDate_to_POSIX(),
          Content              = .x[[8]],
          stringsAsFactors = FALSE))

}

#' COMDate to POSIX
#'
#' @param x COMDate object
#'
#' @return POSIXct
#'
#' @importFrom purrr map_dbl
#' @importFrom openxlsx convertToDateTime
#'
.COMDate_to_POSIX <- function(x) {

  stopifnot('COMDate' %in% class(x))

  x %>% purrr::map_dbl( ~ .x) %>% openxlsx::convertToDateTime()

}