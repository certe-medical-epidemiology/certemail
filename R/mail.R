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

#' Send emails using Microsoft 365
#'
#' This uses the `Microsoft365R` package to send email via Microsoft 365. Connection will be made internally using [Microsoft365R::get_business_outlook()] using the Certe tenant ID.
#' @param body body of email, allows markdown if `markdown = TRUE`
#' @param subject subject of email
#' @param to field 'to', can be character vector
#' @param attachment character (vector) of file location(s)
#' @param header,footer extra text for header or footer, allows markdown if `markdown = TRUE`. See [blastula::blocks()] to build blocks for these sections.
#' @param background background of the surrounding area in the email. Use `""`, `NULL` or `FALSE` to remove background.
#' @param send directly send email, `FALSE` will show the email in the Viewer pane and will ask whether the email should be saved to the Drafts folder of the current Microsoft 365 user.
#' @param cc field 'CC', can be character vector
#' @param bcc field 'BCC', can be character vector
#' @param reply_to field 'reply-to'
#' @param markdown treat body, header and footer as markdown
#' @param signature add signature to email
#' @param signature_name signature name
#' @param signature_address signature email address, placed directly below `signature_name`
#' @param automated_notice print a notice that the mail was send automatically (default is `TRUE` is not in [interactive()] mode)
#' @param ... arguments for [mail()]
#' @param save_location location to save email object to, which consists of all email details and can be printed in the R console
#' @param expect expression which should return `TRUE` prior to sending the email
#' @details [mail_on_error()] can be used for automated scripts.
#'
#' [mail_plain()] send a plain email, without markdown support and with no signature.
#' @rdname mail
#' @importFrom certestyle colourpicker format2
#' @importFrom blastula compose_email md add_attachment
#' @importFrom htmltools HTML
#' @importFrom magrittr `%>%`
#' @export
mail <- function(body,
                 subject,
                 to,
                 cc = NULL,
                 bcc = read_secret("mail.auto_bcc"),
                 reply_to = NULL,
                 attachment = NULL,
                 header = FALSE,
                 footer = FALSE,
                 background = colourpicker("certeblauw6"),
                 send = TRUE,
                 markdown = TRUE,
                 signature = TRUE,
                 signature_name = get_name_and_job_title(),
                 signature_address = get_mail_address(),
                 automated_notice = !interactive(),
                 save_location = read_secret("mail.export_path"),
                 expect = NULL,
                 ...) {

  expect_deparsed <- deparse(substitute(expect))
  if (expect_deparsed != "NULL") {
    expect_check <- tryCatch(expect, error = function(e) FALSE)
    if (!isTRUE(expect_check)) {
      mail_on_error(stop("Mail expection not met: `", expect_deparsed, "`"))
      return(invisible())
    }
  }

  o365 <- get_outlook365(error_on_fail = TRUE)

  # to support HTML
  body <- gsub("<br>", "\n", body, fixed = TRUE)

  if (is.null(background) || background %in% c("", NA, FALSE)) {
    background <- "white"
  } else {
    background <- colourpicker(background)
  }
  if (tryCatch(is.null(save_location) || save_location %in% c("", NA, FALSE), error = function(e) TRUE)) {
    save_location <- NULL
  } else {
    save_location <- gsub("\\", "/", save_location, fixed = TRUE)
    save_location <- gsub("/$", "", save_location)
  }

  # Main text ----
  body <- ifelse(is.null(body), "", body)
  signature_address <- validate_mail_address(signature_address)
  if (markdown == TRUE) {
    if (automated_notice == TRUE) {
      body <- paste0(body, "\n\n*Deze mail is geautomatiseerd verstuurd.*", collapse = "")
    }
    if (signature == TRUE) {
      img <- paste0('<div style="font-family: \'Arial Black\', \'Calibri\', \'Verdana\' !important; font-weight: bold !important; color: ',
                    colourpicker("certeblauw"),
                    ' !important; font-size: 16px !important;" class="certelogo">CERTE</div>')
      body <- paste0(c(body,
                       "\n\nMet vriendelijke groet,\n\n",
                       '<div class="white"></div>\n\n',
                       signature_name, "  \n",
                       "[", signature_address, "](mailto:", signature_address, ")\n\n",
                       img,
                       '<p style="font-family: Calibri, Verdana !important; margin-top: 0px !important;">Postbus 909 | 9700 AX Groningen | <a href="https://www.certe.nl">certe.nl</a></p>'),
                     collapse = "")
    }

  } else {
    # no HTML
    background <- "white"
    if (automated_notice == TRUE) {
      body <- paste0(body, "\n\nDeze mail is geautomatiseerd verstuurd.", collapse = "")
    }
    if (signature == TRUE) {
      body <- paste0(body, "\n\nMet vriendelijke groet,\n\n\n",
                     signature_name,
                     "\n", reply_to,
                     "\n\nCERTE",
                     "\nPostbus 909 | 9700 AX Groningen | certe.nl",
                     collapse = "")
    }
  }

  if (markdown == FALSE) {
    markup <- function(x) ifelse(is.logical(x), "", x)
  } else {
    markup <- function(x) md(as.character(ifelse(is.logical(x), "", x)))
  }
  mail_lst <- compose_email(body = markup(body),
                            header = markup(header),
                            footer = markup(footer))

  # edit background
  mail_lst$html_str <- gsub("#f6f6f6", background, mail_lst$html_str, fixed = TRUE)

  if (isFALSE(header)) {
    mail_lst$html_str <- gsub('class="header" style="', 'class="header" style="display: none !important; ', mail_lst$html_str, fixed = TRUE)
  }
  if (isFALSE(footer)) {
    mail_lst$html_str <- gsub('class="footer" style="', 'class="footer" style="display: none !important; ', mail_lst$html_str, fixed = TRUE)
  }

  # attachments
  if (!is.null(attachment)) {
    for (i in 1:length(attachment)) {
      if (!file.exists(attachment[i])) {
        stop("Attachment does not exist: ", attachment[i], call. = FALSE)
      }
      if (Sys.info()['sysname'] == "Windows") {
        attachment[i] <- gsub("/", "\\\\", attachment[i])
      }
      mail_lst <- mail_lst %>% add_attachment(attachment[i])
    }
  }

  # Set Certe theme
  # font
  mail_lst$html_str <- gsub("Helvetica( !important)?", "Calibri,Verdana !important", mail_lst$html_str)
  # remove headers (also under @media)
  mail_lst$html_str <- gsub("h[123] [{].*[}]+?", "", mail_lst$html_str)
  # re-add headers
  mail_lst$html_str <- gsub('(<style media="all" type="text/css">)',
                            paste("\\1",
                                  # needed for desktop Outlook:
                                  "h1,h2,h3,h4,h5,h6,p,div,li,ul,table,span,header,footer {",
                                  "font-family: 'Calibri', 'Verdana' !important;",
                                  "}",
                                  # the rest:
                                  "h1 {",
                                  "font-size: 18px !important;",
                                  "font-weight: bold !important;",
                                  "margin-top: 10px !important;",
                                  "margin-bottom: 0px !important;",
                                  paste0("color: ", colourpicker("certeblauw"), " !important;"),
                                  "}",
                                  "h2 {",
                                  "font-size: 16px !important;",
                                  "font-weight: bold !important;",
                                  "margin-top: 10px !important;",
                                  "margin-bottom: 0px !important;",
                                  paste0("color: ", colourpicker("certeblauw"), " !important;"),
                                  "}",
                                  "h3 {",
                                  "font-size: 14px !important;",
                                  "font-weight: bold !important;",
                                  "margin-top: 10px !important;",
                                  "margin-bottom: 0px !important;",
                                  paste0("color: ", colourpicker("certeroze"), " !important;"),
                                  "}",
                                  "h4 {",
                                  "font-size: inherit !important;",
                                  "font-weight: bold !important;",
                                  "color: black !important;",
                                  "}",
                                  "p {",
                                  "margin-bottom: 10px;", # no !important since tables contain p too
                                  "margin-top: 0px;",
                                  "}",
                                  "table,img {",
                                  "margin-bottom: 15px !important;",
                                  "}",
                                  # also code, for `text`
                                  "code {",
                                  paste0("color: ", colourpicker("certeroze"), " !important;"),
                                  paste0("background-color: ", colourpicker("certeroze3"), " !important;"),
                                  "font-family: 'Fira Code', 'Courier New' !important;",
                                  'font-size: 12px !important;',
                                  "padding-left: 3px !important;",
                                  "padding-right: 3px !important;",
                                  "padding-top: 2px !important;",
                                  "padding-bottom: 3px !important;",
                                  "}",
                                  # logo for email
                                  ".mail_logo {",
                                  "margin-top: 5px !important;",
                                  "margin-bottom: -5px !important;",
                                  "}",
                                  # the dot after 'Met vriendelijke groet' for extra space
                                  "div.white, .white {",
                                  "height: 18px !important;",
                                  "}",
                                  ".certelogo{",
                                  paste0("color: ", colourpicker("certeblauw"), " !important;"),
                                  "font-family: 'Arial Black', 'Calibri', 'Verdana' !important;",
                                  "font-weight: bold !important;",
                                  "font-size: 16px !important;",
                                  "}",
                                  sep = "\n"),
                            mail_lst$html_str)
  # For old Outlook EXTRA force of style
  mail_lst$html_str <- gsub('<(h[1-6]|p|div|li|ul|table|span|header|footer)>',
                            '<\\1 style="font-family: Calibri, Verdana !important;">', mail_lst$html_str)
  # html element in list in right structure
  mail_lst$html_html <- HTML(mail_lst$html_str)


  subject <- ifelse(is.null(subject), "", trimws(subject))

  reply_to <- validate_mail_address(reply_to)
  to <- validate_mail_address(to)
  cc <- validate_mail_address(cc)
  if (!is.null(bcc)) {
    # remove addresses reply_to bcc that are already in other fields
    bcc <- bcc[!bcc %in% c(to, cc)]
  }
  bcc <- validate_mail_address(bcc)

  actual_mail <- o365$create_email(mail_lst,
                                   to = to,
                                   cc = cc,
                                   bcc = bcc,
                                   reply_to = unname(reply_to),
                                   subject = subject)
  actual_mail_out <- structure(mail_lst,
                               class = c("certe_mail", class(mail_lst)),
                               to = to,
                               cc = cc,
                               bcc = bcc,
                               reply_to = reply_to,
                               subject = subject,
                               attachment = attachment,
                               date_time = Sys.time())

  if (send == TRUE) {
    actual_mail$send()
    message("Mail sent (using Microsoft 365, ", o365$properties$mail, ") at ", format(Sys.time()),
            " with subject '", subject, "'",
            " reply to ", reply_to,
            " to ", concat(to, ", "),
            ifelse(length(cc) > 0,
                   paste0(" and CC'd to ", concat(cc, ", ")),
                   ""),
            ifelse(length(bcc) > 0,
                   paste0(" and BCC'd to ", concat(bcc, ", ")),
                   ""),
            ".")
    if (!is.null(save_location)) {
      # save email object
      filename <- paste0("mail_", format(Sys.time(), "%Y%m%d_%H%M%S"), "_", gsub("[^a-zA-Z0-9]", "_", tolower(subject)), ".rds")
      saveRDS(actual_mail_out, file = paste0(save_location, "/", filename), compress = "xz", version = 2)
    }
    return(invisible())

  } else {
    # not ready to send, save to drafts folder and return object
    if (isTRUE(utils::askYesNo(paste0("Save email to the '", o365$get_drafts()$properties$displayName, "' folder?")))) {
      actual_mail$move(o365$get_drafts())
      message("Concept email saved to '", o365$get_drafts()$properties$displayName, "' folder of account ", o365$properties$mail, ".")
    }
    return(actual_mail_out)
  }
}

#' @rdname mail
#' @export
mail_plain <- function(body,
                       subject,
                       to,
                       cc = NULL,
                       bcc = read_secret("mail.auto_bcc"),
                       reply_to = read_secret("mail.auto_reply_to"),
                       attachment = NULL,
                       send = TRUE,
                       ...) {
  mail(body = body,
       subject = subject,
       to = to,
       cc = cc,
       bcc = bcc,
       reply_to = reply_to,
       attachment = attachment,
       send = send,
       header = FALSE,
       footer = FALSE,
       background = NULL,
       markdown = FALSE,
       signature = FALSE,
       automated_notice = FALSE,
       ...)
}

#' @rdname mail
#' @importFrom blastula add_image
#' @param image_path path of image
#' @param width required width of image, must be in CSS style such as "200px" or "100%"
#' @export
mail_image <- function(image_path, width = NULL, ...) {
  if (is.null(width)) {
    width <- ""
  } else {
    width <- paste0('width="', width, '"')
  }
  if (file.exists(image_path)) {
    img <- add_image(image_path, alt = "") %>%
      gsub('width="520"', width, .) %>%
      gsub("base64", "charset=utf-8;base64", ., fixed = TRUE) %>%  # add encoding
      paste("\n\n", collapse = "")
    if (isTRUE(list(...)$remove_cid)) {
      img <- gsub("cid=.* .*?", "", img)
    }
    img
  } else {
    warning("Image file does not exist: ", normalizePath(image_path), call. = FALSE)
    return("")
  }
}

#' @rdname mail
#' @param expr expression to test, an email will be sent if this expression returns an error
#' @export
mail_on_error <- function(expr, to = read_secret("mail.error_to"), ...) {
  expr_txt <- paste0(deparse(substitute(expr)), collapse = "  \n")
  if (expr_txt %like% "Mail expection not met") {
    expr_txt <- "mail(...)"
  }
  tryCatch(expr = expr,
           error = function(e) {
             call_txt <- trimws(gsub("([/+*^-])", " \\1 ", paste0(deparse(e$call), collapse = "  \n")))
             expr_txt <- trimws(gsub("([/+*^-])", " \\1 ", expr_txt))
             err_text <- paste0(c("Er heeft zich een fout voorgedaan bij het verwerken van de volgende expressie:",
                                  ifelse(call_txt == expr_txt,
                                         paste0("`", expr_txt, "`"),
                                         paste0("`", expr_txt, "`\n\nCall:\n\n`", call_txt, "`")),
                                  paste0("Ingelogde gebruiker: ", unname(Sys.info()["user"])),
                                  paste0("Fout: **", trimws(e$message), "**")),
                                collapse = "\n\n")
             tryCatch(mail(body = err_text,
                           subject = paste0("! Fout bij verwerken expressie (", Sys.info()["user"], ")"),
                           to = to,
                           background = colourpicker("certeroze3"),
                           markdown = TRUE,
                           signature = FALSE,
                           automated_notice = FALSE,
                           send = TRUE,
                           ...),
                      error = function(e) invisible())
             message("Fout bij verwerken van expressie:\n  ", expr_txt,
                     "\nCall:\n  ", call_txt,
                     "\nIngelogde gebruiker:\n  ", unname(Sys.info()["user"]),
                     "\nFout:\n  ", trimws(e$message))})
}

#' @rdname mail
#' @param path location to save attachment(s) in
#' @param search an ODATA filter, ignores `sort`
#' @param search_subject a [character], equal to `search = "subject:(search_subject)"`, case-insensitive
#' @param search_from a [character], equal to `search = "from:(search_from)"`, case-insensitive
#' @param search_when a [Date] vector of size 1 or 2, equal to `search = "received:date1..date2"`, see *Examples*
#' @param folder email folder to search in, defaults to Inbox name of the current user by calling [get_inbox_name()]
#' @param n maximum number of emails to list
#' @param sort initial sorting
#' @param overwrite logical to indicate whether existing local files should be overwritten
#' @details If both `search_subject` and `search_from` are provided, they will be matched as 'AND'. `search_from` can contain any sender name or email address. If `search_when` has a length over 2, the first and last value will be taken.
#' @export
#' @importFrom crayon bold blue
#' @examples
#' \dontrun{
#' download_mail_attachment(search_when = "2021-10-28")
#'
#' if (require("certetoolbox")) {
#'   download_mail_attachment(search_from = "glims", search_when = last_week())
#'
#'   download_mail_attachment(search_from = "drenthe", search_when = yesterday(), n = 1)
#' }
#' }
download_mail_attachment <- function(path = getwd(),
                                     search = NULL,
                                     search_subject = NULL,
                                     search_from = NULL,
                                     search_when = NULL,
                                     folder = get_inbox_name(),
                                     n = 50,
                                     sort = "received desc",
                                     overwrite = TRUE) {
  o365 <- get_outlook365(error_on_fail = TRUE)
  folder <- o365$get_folder(folder)

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
    if (length(search_when) == 1) {
      search_when <- rep(search_when, 2)
    } else if (length(search_when) > 2) {
      search_when <- c(search_when[1], search_when[length(search_when)])
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
                               bold("u"),
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
                            graphics = interactive(),
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
                             graphics = interactive(),
                             title = paste0("Which attachment of mail ", mail_int, " (0 for Cancel)?"))
      if (att_int == 0) {
        return(invisible(NULL))
      }
      att <- att[[att_int]]
    }
    # now download
    if (basename(path) %unlike% "[.]") {
      # no filename yet
      path <- paste0(path, "/", att$properties$name)
    }
    message("Saving file '", att$properties$name, "' as '", path, "'...", appendLF = FALSE)
    att$download(dest = path, overwrite = overwrite)
    message("OK")
  } else {
    # non-interactive, download all attachments
    if (basename(path) %like% "[.]") {
      # has a file name, take top folder
      path <- dirname(path)
    }
    for (i in seq_len(mails)) {
      mail <- mails[[i]]
      att <- mail$list_attachments()
      for (a in seq_len(att)) {
        att_this <- att[[a]]
        p <- paste0(path, "/", att_this$properties$name)
        message("Saving file '", att_this$properties$name, "' as '", path, "'...", appendLF = FALSE)
        att_this$download(dest = p, overwrite = overwrite)
        message("OK")
      }
    }
  }
  return(invisible(path))
}

#' @noRd
#' @method print certe_mail
#' @importFrom crayon bold
#' @importFrom certestyle format2
#' @export
print.certe_mail <- function (x, ...) {
  obj_name <- deparse(substitute(x))
  cat(paste0(bold("Details of email creation:\n"),
             "Date/time: ", format2(attr(x, "date_time", exact = TRUE), "d mmmm yyyy HH:MM:SS"), "\n",
             "Subject:   ", attr(x, "subject", exact = TRUE), "\n",
             "Reply to:  ", ifelse(!is.null(names(attr(x, "reply_to", exact = TRUE))),
                                   paste0("'", names(attr(x, "reply_to", exact = TRUE)), "' <", attr(x, "reply_to", exact = TRUE), ">\n"),
                                   paste0(attr(x, "reply_to", exact = TRUE), "\n")),
             "To:        ", concat(attr(x, "to", exact = TRUE), ", "), "\n",
             "CC:        ", concat(attr(x, "cc", exact = TRUE), ", "), "\n",
             "BCC:       ", concat(attr(x, "bcc", exact = TRUE), ", "), "\n",
             ifelse(length(attr(x, "attachment", exact = TRUE)) == 0,
                    "",
                    paste0("Attachments:\n", paste0(paste0("- ", attr(x, "attachment", exact = TRUE)), collapse = "\n"),
                           "\n"))))
  print(structure(x, class = class(x)[class(x) != "certe_mail"]))
}
