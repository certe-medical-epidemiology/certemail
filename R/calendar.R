# ===================================================================== #
#  An R package by Certe:                                               #
#  https://github.com/certe-medical-epidemiology                        #
#                                                                       #
#  Licensed as GPL-v2.0.                                                #
#                                                                       #
#  Developed at non-profit organisation Certe Medical Diagnostics &     #
#  Advice, department of Medical Epidemiology.                          #
#                                                                       #
#  This R package is free software; you can freely use and distribute   #
#  it for both personal and commercial purposes under the terms of the  #
#  GNU General Public License version 2.0 (GNU GPL-2), as published by  #
#  the Free Software Foundation.                                        #
#                                                                       #
#  We created this package for both routine data analysis and academic  #
#  research and it was publicly released in the hope that it will be    #
#  useful, but it comes WITHOUT ANY WARRANTY OR LIABILITY.              #
# ===================================================================== #

#' Connect to Calendar in Outlook 365
#'
#' After connecting using [connect_outlook()], the `get_*()` functions retrieve properties from this connection.
#' @param account the Microsoft 365 Outlook account object, e.g. as returned by [connect_outlook()]
#' @param start_date,end_date date range
#' @rdname calendar
#' @name calendar
#' @details [calendar_get_events()] retrieves a [data.frame] with all appointments from the shared calendar set in `account`.
#' @importFrom httr GET add_headers content
#' @importFrom dplyr tibble mutate if_else arrange desc
#' @importFrom certeprojects connect_outlook
#' @export
calendar_get_events <- function(start_date = Sys.Date(), end_date = start_date, account = connect_outlook()) {
  my_calendargroups <- GET(url = paste0("https://graph.microsoft.com/v1.0/me/calendargroups"),
                           config = add_headers(Authorization = paste(account$token$credentials$token_type,
                                                                      account$token$credentials$access_token))) |>
    content()
  my_calendargroups_ids <- vapply(FUN.VALUE = character(1), my_calendargroups$value, function(x) x$id)
  calendar_names <- character(0)
  target_calendargroup_id <- ""
  target_calendar_id <- ""
  for (id in my_calendargroups_ids) {
    calendars <- GET(url = paste0("https://graph.microsoft.com/v1.0/me/calendargroups/", id, "/calendars"),
                     config = add_headers(Authorization = paste(account$token$credentials$token_type,
                                                                account$token$credentials$access_token))) |>
      content()
    out <- which(vapply(FUN.VALUE = character(1), calendars$value, function(x) x$owner$address) == account$properties$mail)
    if (length(out) > 0) {
      target_calendargroup_id <- id
      target_calendar_id <- calendars$value[[out]]$id
      break
    }
  }
  if (target_calendar_id == "") {
    stop("Target calendar not found: ", connect_outlook()$properties$mail, ". Is the calendar in your Shared Calendars?")
  }

  calendar <- GET(url = paste0("https://graph.microsoft.com/v1.0/me/calendargroups/", target_calendargroup_id,
                               "/calendars/", target_calendar_id),
                  config = add_headers(Authorization = paste(account$token$credentials$token_type,
                                                             account$token$credentials$access_token))) |>
    content()

  message("Calendar: ", calendar$owner$name, " (", calendar$owner$address, ")")

  events <- GET(url = paste0("https://graph.microsoft.com/v1.0/me/calendargroups/", target_calendargroup_id,
                             "/calendars/", target_calendar_id, "/calendarview",
                             "?StartDateTime=", format(as.Date(start_date)), "T00:00:00.0000000",
                             "&EndDateTime=", format(as.Date(end_date)), "T23:59:59.0000000"),
                config = add_headers(Authorization = paste(account$token$credentials$token_type,
                                                           account$token$credentials$access_token))) |>
    content()

  events <- events$value
  events_df <- tibble(
    calendar_name = calendar$owner$name,
    calendar_address = calendar$owner$address,
    subject = vapply(FUN.VALUE = character(1), events, function(x) x$subject),
    start = as.POSIXct(gsub("T", " ", vapply(FUN.VALUE = character(1), events, function(x) x$start$dateTime)), tz = "UTC"),
    end = as.POSIXct(gsub("T", " ", vapply(FUN.VALUE = character(1), events, function(x) x$end$dateTime)), tz = "UTC"),
    is_all_day = vapply(FUN.VALUE = logical(1), events, function(x) x$isAllDay),
    show_as = vapply(FUN.VALUE = character(1), events, function(x) x$showAs)) |>
    mutate(start = as.POSIXct(start, tz = "Europe/Amsterdam"),
           end = as.POSIXct(end, tz = "Europe/Amsterdam"),
           end = if_else(is_all_day, end - 1, end)) |>
    arrange(desc(start))

  events_df
}

#' @rdname calendar
#' @details
#' [calendar_unavailable_users()] returns a vector of appointment subjects that are shown as "free" within `start_date` and `end_date`.
#'
#' @export
calendar_unavailable_users <- function(start_date = Sys.Date(), end_date = start_date, account = connect_outlook()) {
  events_df <- calendar_get_events(start_date = as.Date(start_date), end_date = as.Date(end_date), account = account)
  events_df$subject[events_df$show_as == "free"]
}
