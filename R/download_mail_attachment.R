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

#' Download Email Attachments Using Microsoft 365
#'
#' This uses the `Microsoft365R` package to download email attachments via Microsoft 365. Connection will be made internally using [`get_business_outlook()`][Microsoft365R::get_business_outlook()] using the Certe tenant ID.
#' @param path location to save attachment(s) in
#' @param filename new filename for the attachments, use `NULL` to not alter the filename. The following texts can be used for replacement (invalid filename characters will be replaced with an underscore):
#'
#' * `"{date}"`, the received date of the email in format "yyyymmdd"
#' * `"{time}"`, the received time of the email in format "HHMMSS"
#' * `"{datetime}"` and `"{date_time}"`, the received date and time of the email in format "yyyymmdd_HHMMSS"
#' * `"{name}"`, the name of the sender of the email
#' * `"{address}"`, the email address of the sender of the email
#' * `"{original}"`, the original filename of the attachment
#' @param search an ODATA filter, ignores `sort` and defaults to search only mails with attachments
#' @param search_subject a [character], equal to `search = "subject:(search_subject)"`, case-insensitive
#' @param search_from a [character], equal to `search = "from:(search_from)"`, case-insensitive
#' @param search_when a [Date] vector of size 1 or 2, equal to `search = "received:date1..date2"`, see *Examples*
#' @param folder email folder to search in, defaults to Inbox name of the current user by calling [get_inbox_name()]
#' @param n maximum number of emails to choose from
#' @param sort initial sorting
#' @param overwrite logical to indicate whether existing local files should be overwritten
#' @param account a Microsoft 365 account to use for sending the mail. This has to be an object as returned by [connect_outlook365()] or [Microsoft365R::get_business_outlook()].
#' @details `search_*` arguments will be matched as 'AND'.
#'
#' `search_from` can contain any sender name or email address. If `search_when` has a length over 2, the first and last value will be taken.
#'
#' See this page for all filtering options using the `search` argument: <https://support.microsoft.com/en-gb/office/how-to-search-in-outlook-d824d1e9-a255-4c8a-8553-276fb895a8da>.
#' @seealso [mail()]
#' @export
#' @importFrom crayon bold blue
#' @examples
#' \dontrun{
#' download_mail_attachment(search_when = "2021-10-28")
#'
#' if (require("certetoolbox")) {
#'   download_mail_attachment(search_from = "groningen",
#'                            search_when = last_week(),
#'                            filename = "groningen_{date_time}")
#'
#'   download_mail_attachment(search_from = "drenthe",
#'                            search_when = yesterday(),
#'                            n = 1)
#' }
#' }
download_mail_attachment <- function(path = getwd(),
                                     filename = "mail_{name}_{datetime}",
                                     search = "hasattachment:yes",
                                     search_subject = NULL,
                                     search_from = NULL,
                                     search_when = NULL,
                                     folder = get_inbox_name(account = account),
                                     n = 5,
                                     sort = "received desc",
                                     overwrite = TRUE,
                                     account = connect_outlook365()) {
  if (!is_valid_o365(account)) {
    message("No valid Microsoft 365 account set with argument `account`")
    return(invisible())
  }

  folder <- account$get_folder(folder)

  if (!is.null(search_subject)) {
    search <- paste0(search, " AND subject:(", search_subject, ")")
  }
  if (!is.null(search_from)) {
    search <- paste0(search, " AND from:(", search_from, ")")
  }
  if (!is.null(search_when)) {
    if (!inherits(search_when, c("Date", "POSIXt")) && any(search_when %unlike% "[0-9]{4}-[0-9]{2}-[0-9]{2}")) {
      stop("`search_when` must be of class Date", call. = FALSE)
    }
    search_when <- as.Date(search_when) # force in case of POSIXt
    if (length(search_when) == 2) {
      # be sure to put oldest first, and latest last
      search_when <- sort(search_when)
    } else if (length(search_when) == 1) {
      search_when <- rep(search_when, 2)
    } else {
      # take oldest and latest date
      search_when <- c(min(search_when), max(search_when))
    }
    search <- paste0(search, " AND received:", search_when[1], "..", search_when[2])
  }
  if (!is.null(search)) {
    if (sort != "received desc") {
      # "received desc" is the default
      message("'search' provided, ignoring 'sort = ", sort, "'")
    }
    search <- gsub("^ AND ", "", search)
  }

  mails <- folder$list_emails(n = n, by = sort, search = search)
  has_attachment <- sapply(mails,
                           function(m) {
                             m$properties$hasAttachments &&
                               sapply(m$list_attachments(), function(a) a$properties$name != "0")
                           })
  if (!any(has_attachment)) {
    if (!is.null(search)) {
      warning("No mails with attachment found in the ", n, " mails searched (search = \"", search, "\")", call. = FALSE)
    } else {
      warning("No mails with attachment found in the ", n, " mails searched (sort = \"", sort, "\")", call. = FALSE)
    }
    return(NULL)
  }
  mails <- mails[which(has_attachment)]
  mails_txt <- sapply(mails,
                      function(m) {
                        p <- m$properties
                        dt <- as.POSIXct(gsub("T", " ", p$receivedDateTime), tz = "UTC")
                        paste0(bold(format2(dt, "ddd d mmm yyyy HH:MM")),
                               "\n    ", p$subject, " (",
                               ifelse(p$from$emailAddress$name == p$from$emailAddress$address,
                                      blue(paste0(p$from$emailAddress$address)),
                                      blue(paste0(p$from$emailAddress$name, ", ", p$from$emailAddress$address))),
                               ")\n",
                               paste0("    - ", sapply(m$list_attachments(), function(a) {
                                 if (a$properties$size > 1024 ^ 2) {
                                   size <- paste0(round(a$properties$size / 1024 ^ 2, 1), " MB")
                                 } else if (a$properties$size > 1024) {
                                   size <- paste0(round(a$properties$size / 1024, 0), " kB")
                                 } else {
                                   size <- paste0(a$properties$size, " B")
                                 }
                                 paste0(a$properties$name, " (", size , ")")
                               }),
                               collapse = "\n")
                        )
                      })
  if (interactive()) {
    # pick mail
    mail_int <- utils::menu(mails_txt,
                            graphics = FALSE,
                            title = paste0(length(mails),
                                           " mails found with attachment. Which mail to select (0 for Cancel)?"))
    if (mail_int == 0) {
      return(invisible(NULL))
    }
    mail <- mails[[mail_int]]
    att <- mail$list_attachments()
    if (length(att) == 1) {
      att <- att[[1]]
    } else {
      # pick attachment
      att_int <- utils::menu(sapply(att, function(a) paste0(a$properties$name, " (", round(a$properties$size / 1024), " kB)")),
                             graphics = FALSE,
                             title = paste0("Which attachment of mail ", mail_int, " (0 for Cancel)?"))
      if (att_int == 0) {
        return(invisible(NULL))
      }
      att <- att[[att_int]]
    }
    # now download
    if (basename(path) %unlike% "[.]") {
      # no filename yet
      path <- paste0(path, "/", format_filename(mail, att, filename))
    }
    message("Saving file '", att$properties$name, "' as '", path, "'...", appendLF = FALSE)
    tryCatch({
      att$download(dest = path, overwrite = overwrite)
      message("OK")
    }, error = function(e) {
      message("ERROR:\n", e$message)
    })
  } else {
    # non-interactive, download all attachments
    if (basename(path) %like% "[.]") {
      # has a file name, take top folder
      path <- dirname(path)
    }
    for (i in seq_len(length(mails))) {
      mail <- mails[[i]]
      att <- mail$list_attachments()
      for (a in seq_len(length(att))) {
        att_this <- att[[a]]
        p <- paste0(path, "/", format_filename(mail, att_this, filename))
        message("Saving file '", att_this$properties$name, "' as '", p, "'...", appendLF = FALSE)
        tryCatch({
          att_this$download(dest = p, overwrite = overwrite)
          message("OK")
        }, error = function(e) {
          message("ERROR:\n", e$message)
        })
      }
    }
  }
  return(invisible(path))
}

#' @importFrom certestyle format2
format_filename <- function(mail, attachment, filename) {
  if (is.null(filename)) {
    return(attachment$properties$name)
  }

  dt <- as.POSIXct(gsub("T", " ", mail$properties$receivedDateTime), tz = "UTC")
  # replace text items with relative properties
  filename <- gsub("{date}", format2(dt, "yyyymmdd"), filename, fixed = TRUE)
  filename <- gsub("{time}", format2(dt, "HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{datetime}", format2(dt, "yyyymmdd_HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{date_time}", format2(dt, "yyyymmdd_HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{name}", mail$properties$from$emailAddress$name, filename, fixed = TRUE)
  filename <- gsub("{address}", mail$properties$from$emailAddress$address, filename, fixed = TRUE)
  filename <- gsub("{original}", attachment$properties$name, filename, fixed = TRUE)

  # replace invalid filename characters
  filename <- trimws(gsub("[\\/:*?\"<>|]+", "_", filename))

  # add extension if missing
  ext <- gsub(".*([.].*)$", "\\1", attachment$properties$name)
  if (filename %unlike% paste0(ext, "$")) {
    filename <- paste0(filename, ext)
  }

  filename
}
