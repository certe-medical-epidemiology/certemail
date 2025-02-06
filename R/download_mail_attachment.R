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
#' This uses the `Microsoft365R` package to download email attachments via Microsoft 365. Connection will be made using [connect_outlook()].
#' @param path location to save attachment(s) to
#' @param filename new filename for the attachments, use `NULL` to not alter the filename. For multiple files, the files will be appended with a number. The following texts can be used for replacement (invalid filename characters will be replaced with an underscore):
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
#' @param search_attachment a [character] to use a regular expression for attachment file names
#' @param folder email folder name to search in, defaults to Inbox of the current user by calling [get_inbox_name()]
#' @param n maximum number of emails to search
#' @param sort initial sorting
#' @param overwrite a [logical] to indicate whether existing local files should be overwritten
#' @param skip_inline a [logical] to indicate whether inline attachments such as meta images must be skipped
#' @param account a Microsoft 365 account to use for searching the mails. This has to be an object as returned by [connect_outlook()] or [Microsoft365R::get_business_outlook()].
#' @param interactive_mode a [logical] to indicate interactive mode. In non-interactive mode, all attachments within the search criteria will be downloaded.
#' @details `search_*` arguments will be matched as 'AND'.
#'
#' `search_from` can contain any sender name or email address. If `search_when` has a length over 2, the first and last value will be taken.
#'
#' Refer to [this Microsoft Support page](https://support.microsoft.com/en-gb/office/how-to-search-in-outlook-d824d1e9-a255-4c8a-8553-276fb895a8da) for all filtering options using the `search` argument.
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
                                     search_attachment = NULL,
                                     folder = get_inbox_name(account = account),
                                     n = 5,
                                     sort = "received desc",
                                     overwrite = TRUE,
                                     skip_inline = TRUE,
                                     account = connect_outlook(),
                                     interactive_mode = interactive()) {
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
    search_when <- as.Date(search_when) # force in case of POSIXt and character
    if (length(search_when) == 2) {
      # be sure to put oldest first, and latest last
      search_when <- sort(search_when)
    } else if (length(search_when) == 1) {
      search_when <- rep(search_when, 2)
    } else {
      # take oldest and newest date
      search_when <- c(min(search_when, na.rm = TRUE), max(search_when, na.rm = TRUE))
    }
    search <- paste0(search, " AND received:", search_when[1], "..", search_when[2])
  }
  if (!is.null(search)) {
    if (sort != "received desc") {
      # sort cannot be used if search is being used
      message("'search' provided, ignoring 'sort = ", sort, "'")
    }
    search <- gsub("^ AND ", "", search)
  }

  mails <- folder$list_emails(n = n, by = sort, search = search)
  message(length(mails), " mail", ifelse(length(mails) == 1, "", "s"), " found with filter '", search, "'")
  has_attachment <- vapply(FUN.VALUE = logical(1),
                           mails,
                           function(m) {
                             m$properties$hasAttachments &
                               any(vapply(FUN.VALUE = logical(1),
                                          only_valid_attachments(m$list_attachments(),
                                                                 search = search_attachment,
                                                                 skip_inline = skip_inline),
                                          function(att) att$properties$name != "0" && ifelse(skip_inline, !att$properties$isInline, TRUE)))
                           })

  if (!any(has_attachment)) {
    if (!is.null(search) || !is.null(search_attachment)) {
      warning("No mails with attachment found in the ",
              n, " mails searched (search = \"", search, "\")",
              ifelse(is.null(search_attachment),
                     "", paste0(" and search_attachment = \"", search_attachment, "\"")),
              call. = FALSE)
    } else {
      warning("No mails with attachment found in the ", n, " mails searched (sort = \"", sort, "\")", call. = FALSE)
    }
    return(NULL)
  }
  mails <- mails[which(has_attachment)]
  mails_txt <- vapply(FUN.VALUE = character(1),
                      mails,
                      function(m) {
                        p <- m$properties
                        # assume the mails were received in our timezone
                        dt <- as.POSIXct(format(as.POSIXct(gsub("T", " ", p$receivedDateTime),
                                                           tz = "UTC"),
                                                tz = "Europe/Amsterdam"))
                        paste0(bold(format2(dt, "ddd d mmm yyyy HH:MM")),
                               "\n    ", p$subject, " (",
                               ifelse(p$from$emailAddress$name == p$from$emailAddress$address,
                                      blue(paste0(p$from$emailAddress$address)),
                                      blue(paste0(p$from$emailAddress$name, ", ", p$from$emailAddress$address))),
                               ")\n",
                               "      Attachments:\n",
                               paste0(vapply(FUN.VALUE = character(1),
                                             only_valid_attachments(m$list_attachments(),
                                                                    search = search_attachment,
                                                                    skip_inline = skip_inline),
                                             function(a) paste0("      - ", a$properties$name, " (", size_formatted(a$properties$size) , ")\n")),
                                      collapse = ""),
                               ifelse(skip_inline, "      (inline attachments not shown, since `skip_inline = TRUE`)", "")
                        )
                      })

  if (isTRUE(interactive_mode)) {
    if (length(mails) > 1) {
      # pick mail
      mail_int <- utils::menu(c("All", mails_txt),
                              graphics = FALSE,
                              title = "Which mail to select (0 for Cancel)?")
      if (mail_int == 0) {
        return(invisible(NULL))
      } else if (mail_int == 1) {
        # chose 'All'
        mail_object <- mails
      } else {
        mail_object <- mails[[mail_int - 1]] # since 1 was added as 'All'
      }
    } else {
      mail_object <- mails[[1]]
    }
  }

  if (isTRUE(interactive_mode) && inherits(mail_object, "ms_outlook_email")) {
    # is a single mail
    dt <- as.POSIXct(format(as.POSIXct(gsub("T", " ", mail_object$properties$receivedDateTime),
                                       tz = "UTC"),
                            tz = "Europe/Amsterdam"))
    att <- only_valid_attachments(mail_object$list_attachments(),
                                  search = search_attachment,
                                  skip_inline = skip_inline)
    if (length(att) > 1) {
      # pick attachment
      att_int <- utils::menu(c("All",
                               vapply(FUN.VALUE = character(1),
                                      att,
                                      function(a) paste0(a$properties$name, " (", size_formatted(a$properties$size), ")"))),
                             graphics = FALSE,
                             title = paste0("Which attachment of mail ", mail_int, ", from ", mail_object$properties$from$emailAddress$name, " (0 for Cancel)?"))
      if (att_int == 0) {
        return(invisible(NULL))
      }
      if (att_int == 1) {
        # chose All
        att_int <- seq_len(length(att))
      } else {
        att_int <- att_int - 1 # since All is the first
      }
    } else {
      att_int <- 1
    }
    for (i in att_int) {
      att_current <- att[[i]]
      # now download
      if (basename(path) %unlike% "[.]") {
        # no filename yet
        path <- paste0(path, "/", format_filename(mail_object, att_current, filename, add_seq_if_exists = TRUE))
      }
      message("Saving attachment '", att_current$properties$name,
              "' from ", mail_object$properties$from$emailAddress$address,
              "'s email of ", format2(dt, "ddd d mmm yyyy HH:MM") ,
              " to '", path, "'...",
              appendLF = FALSE)
      tryCatch({
        att_current$download(dest = path, overwrite = overwrite)
        message("OK")
      }, error = function(e) {
        message("ERROR:\n", e$message)
      })
    }
  } else {
    # non-interactive or chosen for "All", so download all attachments for all mails
    if (basename(path) %like% "[.]") {
      # has a file name, take parent folder
      path <- dirname(path)
    }
    for (i in seq_len(length(mails))) {
      mail_object <- mails[[i]]
      dt <- as.POSIXct(format(as.POSIXct(gsub("T", " ", mail_object$properties$receivedDateTime),
                                         tz = "UTC"),
                              tz = "Europe/Amsterdam"))
      att <- only_valid_attachments(mail_object$list_attachments(),
                                    search = search_attachment,
                                    skip_inline = skip_inline)
      for (a in seq_len(length(att))) {
        att_this <- att[[a]]
        p <- paste0(path, "/", format_filename(mail_object, att_this, filename, add_seq_if_exists = TRUE))
        message("Saving attachment '", att_this$properties$name,
                "' from ", mail_object$properties$from$emailAddress$address,
                "'s email of ", format2(dt, "ddd d mmm yyyy HH:MM") ,
                " to '", p, "'...",
                appendLF = FALSE)
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
