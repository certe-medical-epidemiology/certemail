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

#' Send Emails Using Microsoft 365
#'
#' These functions use the `Microsoft365R` R package to send emails via Microsoft 365. They require an Outlook Business account.
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
#' @param signature text to print as email signature, or `NULL` to omit it, defaults to [get_certe_signature()]
#' @param automated_notice a [logical] to print a notice that the mail was sent automatically (default is `TRUE` if not in [interactive()] mode)
#' @param save_location location to save email object to, which consists of all email details and can be printed in the R console
#' @param sent_subfolder mail folder within Sent Items in the Microsoft 365 account, to store the mail if `!interactive()`
#' @param expect expression which should return `TRUE` prior to sending the email
#' @param account a Microsoft 365 account to use for sending the mail. This has to be an object as returned by [connect_outlook()] or [Microsoft365R::get_business_outlook()]. Using `account = FALSE` is equal to setting `send = FALSE`.
#' @param identifier a mail identifier to be printed at the bottom of the email. Defaults to [`project_identifier()`][certeprojects::project_identifier()]. Use `FALSE` to not print an identifier.
#' @param ... arguments for [mail()]
#' @details [mail_on_error()] can be used for automated scripts.
#'
#' [mail_plain()] sends a plain email, without markdown support and with no signature.
#' @rdname mail
#' @importFrom certestyle colourpicker format2 plain_html_table
#' @importFrom certeprojects connect_outlook
#' @importFrom blastula compose_email md add_attachment
#' @importFrom htmltools HTML
#' @seealso [download_mail_attachment()]
#' @export
#' @examples
#' mail("test123", "test456", to = "mail@domain.com", account = NULL)
#'
#' mail_plain("test123", "test456", to = "mail@domain.com", account = NULL)
#'
#' mail(mail_image(image_path = system.file("test.jpg", package = "certemail")),
#'     "test456", to = "mail@domain.com", account = NULL)
#'
#' # data.frames will be transformed with certestyle::plain_html_table()
#' mail(mtcars[1:5, ],
#'      subject = "Check these cars!",
#'      to = "somebody@domain.org",
#'      account = FALSE)
mail <- function(body,
                 subject = "",
                 to = NULL,
                 cc = read_secret("mail.auto_cc"),
                 bcc = read_secret("mail.auto_bcc"),
                 reply_to = NULL,
                 attachment = NULL,
                 header = FALSE,
                 footer = FALSE,
                 background = certestyle::colourpicker("certeblauw6"),
                 send = TRUE,
                 markdown = TRUE,
                 signature = get_certe_signature(account = account),
                 automated_notice = !interactive(),
                 save_location = read_secret("mail.export_path"),
                 sent_subfolder = read_secret("mail.sent_subfolder"),
                 expect = NULL,
                 account = connect_outlook(),
                 identifier = NULL,
                 ...) {

  expect_deparsed <- deparse(substitute(expect))
  if (expect_deparsed != "NULL") {
    expect_check <- tryCatch(expect, error = function(e) FALSE)
    if (!isTRUE(expect_check)) {
      mail_on_error(stop("Mail expection not met: `", expect_deparsed, "`"))
      return(invisible())
    }
  }

  if (isTRUE(send) && !is_valid_o365(account)) {
    if (!isFALSE(account)) {
      message("No valid Microsoft 365 account set with argument `account`, forcing `send = FALSE`")
    }
    send <- FALSE
  }

  # to support HTML
  if (is.data.frame(body)) {
    body <- plain_html_table(body)
  }
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
  if (isTRUE(automated_notice)) {
    if (isTRUE(markdown)) {
      body <- paste0(body, "\n\n*Deze mail is geautomatiseerd verstuurd.*", collapse = "")
    } else {
      body <- paste0(body, "\n\nDeze mail is geautomatiseerd verstuurd.", collapse = "")
    }
  }

  if (!is.null(signature) && !isFALSE(signature)) {
    body <- paste0(body, "\n\n", signature)
  }

  # add identifier to mail
  if (missing(identifier) && "certeprojects" %in% rownames(utils::installed.packages())) {
    proj <- certeprojects::project_get_current_id(ask = FALSE)
    if (!is.null(proj)) {
      body <- paste0(body,
                     "\n\n<p class='project-identifier'>",
                     certeprojects::project_identifier(project_number = proj),
                     "</p>")
    }
  } else if (!isFALSE(identifier)) {
    body <- paste0(body, "\n\n<p class='project-identifier'>", identifier, "</p>")
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
    for (i in seq_len(length(attachment))) {
      if (!file.exists(attachment[i])) {
        stop("Attachment does not exist: ", attachment[i], call. = FALSE)
      }
      if (Sys.info()["sysname"] == "Windows") {
        attachment[i] <- gsub("/", "\\\\", attachment[i])
      }
      mail_lst <- add_attachment(mail_lst, attachment[i])
    }
  }

  # Set Certe theme ----
  # font
  mail_lst$html_str <- gsub("Helvetica( !important)?", "Calibri,Verdana !important", mail_lst$html_str)
  # remove headers (also under @media)
  mail_lst$html_str <- gsub("h[123] [{].*[}]+?", "", mail_lst$html_str)
  # re-add headers
  mail_lst$html_str <- gsub('(<style(.*?)>)',
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
                                  # for project identifier
                                  ".project-identifier {",
                                  "font-size: 9px !important;",
                                  "font-weight: normal !important;",
                                  "color: #CBCBCB !important;",
                                  "text-align: right !important;",
                                  "}",
                                  # also code, for `text`
                                  "code {",
                                  paste0("color: ", colourpicker("certeroze0"), " !important;"),
                                  paste0("background-color: ", colourpicker("certeroze6"), " !important;"),
                                  "font-family: 'Fira Code', 'Courier New' !important;",
                                  "font-size: 12px !important;",
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
  mail_lst$html_str <- gsub("<(h[1-6]|p|div|li|ul|table|span|header|footer)>",
                            '<\\1 style="font-family: Calibri, Verdana !important;">', mail_lst$html_str)

  # html element in list in right structure
  mail_lst$html_html <- HTML(mail_lst$html_str)


  subject <- ifelse(is.null(subject), "", trimws(subject))

  reply_to <- validate_mail_address(reply_to)
  to <- validate_mail_address(to)
  cc <- validate_mail_address(cc)
  if (!is.null(bcc)) {
    # remove addresses from bcc that are already in other fields
    bcc <- bcc[!bcc %in% c(to, cc)]
  }
  bcc <- validate_mail_address(bcc)

  if (is_valid_o365(account)) {
    actual_mail <- account$create_email(mail_lst,
                                        to = to,
                                        cc = cc,
                                        bcc = bcc,
                                        reply_to = unname(reply_to),
                                        subject = subject)
  }
  actual_mail_out <- structure(mail_lst,
                               class = c("certe_mail", class(mail_lst)),
                               from = get_mail_address(account = account),
                               to = to,
                               cc = cc,
                               bcc = bcc,
                               reply_to = reply_to,
                               body = body,
                               subject = subject,
                               attachment = attachment,
                               date_time = Sys.time())

  if (isTRUE(send)) {
    if (interactive()) {
      print(actual_mail_out, browse_in_viewer = FALSE)
      cat("\n")
      # 'prompts' is required to prevent Windows from showing a popup box instead of asking in the console
      if (!isTRUE(utils::askYesNo("Send the mail?", prompts = c("Yes", "No", "Cancel")))) {
        actual_mail$delete(confirm = FALSE)
        return(invisible())
      }
    }
    actual_mail$send()
    if (!interactive()) {
      message("Mail sent at ", format(Sys.time()), ":")
      print(actual_mail_out, browse_in_viewer = FALSE)
    }
    # move to subfolder if not interactive
    if ((!interactive() || isTRUE(automated_notice)) && !is.null(sent_subfolder) && trimws(sent_subfolder) != "") {
      sent_items <- tryCatch(account$get_folder("sentitems"), error = function(e) NULL)
      if (!is.null(sent_items) && !sent_subfolder %in% vapply(FUN.VALUE = character(1), sent_items$list_folders(), function(x) x$properties$displayName)) {
        # create folder first
        sent_items$create_folder(sent_subfolder)
        message("Created folder '", sent_subfolder, "' within folder '", sent_items$properties$displayName, "'")
      }
      move_try1 <- tryCatch({
        # actual move
        Sys.sleep(2) # this is to prevent a 404 error
        actual_mail$move(sent_items$get_folder(sent_subfolder))
        message("Mail moved to folder '", sent_subfolder, "' within folder '", sent_items$properties$displayName, "'")
        return(TRUE)
      }, error = function(e) {
        message("Mail could not be moved to folder '", sent_subfolder, "' within folder '", sent_items$properties$displayName, "', waiting another 10 seconds...")
        return(FALSE)
      })
      if (move_try1 == FALSE) {
        tryCatch({
          # wait another 10 seconds
          Sys.sleep(10) # this is to prevent a 404 error
          actual_mail$move(sent_items$get_folder(sent_subfolder))
          message("Mail moved to folder '", sent_subfolder, "' within folder '", sent_items$properties$displayName, "'")
        }, error = function(e) {
          warning("Mail could not be moved to folder '", sent_subfolder, "' within folder '", sent_items$properties$displayName, "': ", e$message, call. = FALSE)
        })
      }
    }

    if (!is.null(save_location)) {
      # save email object
      if (!dir.exists(save_location)) {
        warning("Cannot save mail object to unexisting folder '", save_location, "'")
      } else {
        filename <- paste0("mail_", format(Sys.time(), "%Y%m%d_%H%M%S"), "_", gsub("[^a-zA-Z0-9]", "_", tolower(subject)), ".rds")
        saveRDS(actual_mail_out, file = paste0(save_location, "/", filename), compress = "xz", version = 2)
      }
    }
    return(invisible())

  } else {
    # not ready to send, save to drafts folder and return object
    if (is_valid_o365(account) && isTRUE(utils::askYesNo(paste0("Save email to the folder \"", get_drafts_name(account), "\"?")))) {
      actual_mail$move(account$get_drafts())
      message("Draft saved to folder \"", get_drafts_name(account), "\" of account ", get_mail_address(account), ".")
    } else if (is_valid_o365(account)) {
      actual_mail$delete(confirm = FALSE)
    }
    return(actual_mail_out)
  }
}

#' @rdname mail
#' @export
mail_plain <- function(body,
                       subject = "",
                       to = NULL,
                       cc = read_secret("mail.auto_cc"),
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
       signature = NULL,
       header = FALSE,
       footer = FALSE,
       background = NULL,
       markdown = FALSE,
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
    img <- gsub("base64", "charset=utf-8;base64",
                gsub('width="520"', width,
                     add_image(image_path, alt = "")),
                fixed = TRUE) |>  # add encoding
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

  proj <- NULL
  proj_id <- NULL
  if ("certeprojects" %in% rownames(utils::installed.packages())) {
    proj <- certeprojects::project_get_current_id(ask = FALSE)
    proj_id <- certeprojects::project_identifier(project_number = proj)
    if (!is.null(proj)) {
      proj <- paste0("p", proj, " (", certeprojects::project_get_title(proj), ")")
    }
  }

  tryCatch(expr = expr,
           error = function(e) {
             call_txt <- trimws(gsub("([/+*^-])", " \\1 ", paste0(deparse(e$call), collapse = "  \n")))
             full_call_txt <- trimws(gsub("([/+*^-])", " \\1 ", paste0(deparse(sys.calls()), collapse = "  \n")))
             expr_txt <- trimws(gsub("([/+*^-])", " \\1 ", expr_txt))
             err_text <- paste0(c("Mail error:",
                                  ifelse(call_txt == expr_txt,
                                         paste0("`", expr_txt, "`"),
                                         paste0("`", expr_txt, "`\n\nCall:\n\n`", call_txt, "`")),
                                  ifelse(is.null(proj),
                                         "",
                                         paste0("Project:\n\n", proj)),
                                  paste0("User: ", unname(Sys.info()["user"])),
                                  paste0("Error message: **", trimws(e$message), "**")),
                                collapse = "\n\n")
             tryCatch(mail(body = err_text,
                           subject = paste0("! Mail error (", Sys.info()["user"], ") " , proj),
                           to = to,
                           background = colourpicker("certeroze3"),
                           markdown = TRUE,
                           signature = FALSE,
                           automated_notice = FALSE,
                           send = TRUE,
                           identifier = proj_id,
                           ...),
                      error = function(e) invisible())
             message("Error:\n  ", expr_txt,
                     "\nProject:\n  ", proj,
                     "\nCall:\n  ", call_txt,
                     "\nUser:\n  ", unname(Sys.info()["user"]),
                     "\nError message:\n  ", trimws(e$message))})
}

#' @noRd
#' @method print certe_mail
#' @importFrom crayon bold
#' @importFrom certestyle format2
#' @importFrom xml2 xml_text read_html
#' @export
print.certe_mail <- function (x, browse_in_viewer = TRUE, ...) {
  body <- attr(x, "body", exact = TRUE)
  body <- gsub("(\n|\t|<br>)+", " ", body)
  # keep only the text of markdown links
  body <- gsub("\\[(.*)?\\]\\(.*?\\)", "\\1", body)
  body <- tryCatch(xml_text(read_html(body)), error = function(e) body)
  if (nchar(body) > options()$width - 11) {
    body <- paste0(substr(body, 1, options()$width - 13), "...")
  }
  cat(paste0(bold("Mail Summary\n"),
             "Subject:   ", attr(x, "subject", exact = TRUE), "\n",
             "Body text: ", body, "\n",
             "From:      ", attr(x, "from", exact = TRUE), "\n",
             if (!is.null(attr(x, "reply_to", exact = TRUE)) && attr(x, "reply_to", exact = TRUE) != "") {
               paste0("Reply to:  ",
                      ifelse(!is.null(names(attr(x, "reply_to", exact = TRUE))),
                        paste0("'", names(attr(x, "reply_to", exact = TRUE)), "' <", attr(x, "reply_to", exact = TRUE), ">\n"),
                        paste0(attr(x, "reply_to", exact = TRUE), "\n")))
             },
             "To:        ", paste0(attr(x, "to", exact = TRUE), collapse = ", "), "\n",
             "CC:        ", paste0(attr(x, "cc", exact = TRUE), collapse = ", "), "\n",
             "BCC:       ", paste0(attr(x, "bcc", exact = TRUE), collapse = ", "), "\n",
             "Created:   ", format2(attr(x, "date_time", exact = TRUE), "d mmmm yyyy HH:MM:SS"), "\n",
             ifelse(length(attr(x, "attachment", exact = TRUE)) == 0,
                    "",
                    paste0("Attachments:\n", paste0(paste0("- ", attr(x, "attachment", exact = TRUE)), collapse = "\n"),
                           "\n"))))
  if (isTRUE(browse_in_viewer)) {
    print(structure(x, class = class(x)[class(x) != "certe_mail"]))
  }
}

#' @param project_number Number of a project. Will be used to check the grey identifier in the email.
#' @param date A date, defaults to today. Will be evaluated in [as.Date()]. Can also be of length 2 for a date range.
#' @details Use [mail_is_sent()] to check whether a project email was sent on a certain date from any Sent Items (sub)folder.
#'
#' The function will search for the grey identifier in the email body, which is formatted as `[-yymmdd][0-9]+[-project_number][^0-9a-z]`.
#'
#' The function will return `TRUE` if any email was found, and `FALSE` otherwise. If `TRUE`, the name will contain the date(s) and time(s) of the sent email.
#' @importFrom certeprojects connect_outlook
#' @importFrom certestyle format2
#' @rdname mail
#' @export
mail_is_sent <- function(project_number, date = Sys.Date(), account = connect_outlook()) {
  if (!is_valid_o365(account)) {
    stop("`account` is not a valid Microsoft365 account")
  }
  sent_items <- account$get_sent_items()
  sent_items_subfolders <- sent_items$list_folders()
  date <- format(as.Date(date))
  if (length(date) == 1) {
    date <- c(date, date)
  }
  mails <- sent_items$list_emails(by = "received desc",
                                  search = paste0("received:", date[1], "..", date[2]))
  for (i in seq_len(length(sent_items_subfolders))) {
    extra_mails <- sent_items_subfolders[[i]]$list_emails(by = "received desc",
                                                          search = paste0("received:", date[1], "..", date[2]))
    mails <- c(mails, extra_mails)
  }
  if (length(mails) == 0) {
    return(FALSE)
  }

  found <- mails |>
    vapply(FUN.VALUE = logical(1),
           function(x) x$properties$body$content %like% paste0("[-]", format2(date[1], "yymmdd"), "[0-9]+[-]", project_number, "[^0-9a-z]"))
  if (any(found)) {
    datetimes <- mails[found] |>
      vapply(FUN.VALUE = character(1),
             function(x) format2(as.POSIXct(as.POSIXct(gsub("T", " ", x$properties$sentDateTime),
                                                       tz = "UTC"),
                                            tz = "Europe/Amsterdam"),
                                 "yyyy-mm-dd HH:MM:SS")) |>
      sort() |>
      paste(collapse = ", ")
    return(stats::setNames(TRUE, datetimes))
  } else {
    return(FALSE)
  }
}
