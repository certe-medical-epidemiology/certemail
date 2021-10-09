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
#' @param body Tekst van de e-mail.
#' @param subject Onderwerp van de e-mail. Spaties aan begin en einde worden weggesneden.
#' @param to Geaddresseerde (veld 'Aan') van de e-mail. Kan ook vector zijn. Standaard het e-mailadres zoals gedefinieerd in de omgevingsvariabele van de ingelogde gebruiker.
#' @param attachment Bijlagen: vector van locaties. Als de locatie niet bestaat, wordt er een fout gegeven en wordt de mail niet verstuurd.
#' @param header,footer Extra info voor boven en onder de body. Zie \link{blocks} voor het maken van blokken die hier geplaatst kunnen worden. Bij \code{FALSE} worden deze blokken verborgen.
#' @param background Achtergrond van de mail (standaard: \code{\link{colourpicker}("certeblauw3")}). Gebruik \code{""}, \code{NULL} of \code{FALSE} om geen achtergrond te gebruiken.
#' @param send Direct versturen van de e-mail. Met \code{send = FALSE} wordt de e-mail op het scherm weergegeven. Standaard \code{TRUE}.
#' @param cc Geaddresseerde (veld 'CC') van de e-mail. Kan ook vector zijn.
#' @param bcc Geaddresseerde (veld 'BCC') van de e-mail. Kan ook vector zijn. Adressen die voorkomen in \code{to} of \code{cc} worden eruit gehaald. Als in \code{to} of \code{cc} minimaal 1 data-onderzoeker staat, wordt dit standaard niet naar de data-onderzoekers gestuurd als BCC.
#' @param reply_to Geaddresseerde (veld 'Van') van de e-mail. In Outlook moeten de rechten vastgelegd zijn om vanaf dit account te versturen.
#' @param md_body Tekst van de e-mail als Markdown opmaken (standaard: \code{TRUE}).
#' @param signature Handtekening aan e-mail toevoegen (standaard: \code{TRUE}).
#' @param signature_name Naam die onder 'Met vriendelijke groet' komt te staan.
#' @param signature_address E-mailaderes die onder 'Met vriendelijke groet' komt te staan.
#' @param automated_notice Onderaan de mail deze tekst weergeven: "\emph{Deze mail is geautomatiseerd verstuurd.}" (standaard: \code{TRUE}).
#' @param expr Expressie die gedraaid moet worden
#' @param font.size Standaard is 11. Lettergrootte die doorgegeven wordt aan \code{\link{tbl_flextable}()}.
#' @param ... Voor \code{mail_on_error}: parameters die doorgegeven worden aan \code{mail()}. \cr Voor \code{mail_dataset()}: parameters die doorgegeven worden aan \code{\link{tbl_flextable}()}.
#' @param df Een dataset
#' @param path De locatie van een afbeelding
#' @param width De breedte van een afbeelding. Laat leeg om de oorspronkelijke breedte te gebruiken. Dit kan aantal pixels zijn (\code{width = 100} of \code{width = "100px"}) of een percentage van de breedte van de e-mail (\code{width = "25\%"}).
#' @param outlook Mail opstellen met Outlook (standaard: \code{FALSE}). Anders wordt de mail rechtstreeks via de SMTP-server verzonden.
#' @param save_location De maplocatie waar de mail opgeslagen wordt nadat deze verzonden is. Deze wordt opgeslagen als .rds-bestand (met evt. bijlagen), die in RStudio geopend kan worden en geprint kan worden in de Console voor alle maildetails. De mail kan dan zelfs opnieuw verstuurd worden. Gebruik \code{""}, \code{NULL} of \code{FALSE} om de mail niet op te slaan.
#' @param expect Standaard is leeg. Een expressie die \code{TRUE} moet zijn alvorens te mailen. Als dit \code{FALSE} is om een fout oplevert, wordt \code{mail_on_error()} gedraaid met details over de fout.
#' @details De functie \code{mail_on_error()} is speciaal voor gebruik in geautomatiseerde scripts. Het evalueert de input en verstuurt een mail met de expressie wanneer deze een fout retourneert.
#'
#' De functies \code{mail_dataset()} en \code{mail_image()} zijn om resp. een tabel met \code{\link{tbl_flextable}()} en een afbeelding in de mail te plaatsen.
#' @rdname mail
#' @importFrom blastula compose_email md
#' @importFrom htmltools HTML
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
                 background = colourpicker("certeblauw3"),
                 send = TRUE,
                 md_body = TRUE,
                 signature = TRUE,
                 signature_name = get_name_and_job_title(),
                 signature_address = get_mail_address(),
                 automated_notice = !interactive(),
                 outlook = FALSE,
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
  if (md_body == TRUE) {
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

  if (md_body == FALSE) {
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
      mail_export <- structure(mail_lst,
                               class = c("certetools_mail", class(mail_lst)),
                               to = to,
                               cc = cc,
                               bcc = bcc,
                               reply_to = reply_to,
                               subject = subject,
                               attachment = attachment,
                               date_time = Sys.time())
      filename <- paste0("mail_", format(Sys.time(), "%Y%m%d_%H%M%S"), "_", gsub("[^a-zA-Z0-9]", "_", tolower(subject)), ".rds")
      saveRDS(mail_export, file = paste0(save_location, "/", filename), compress = "xz", version = 2)
    }

  } else {
    # not ready to send, save to drafts folder and return object
    actual_mail$move(o365$get_drafts())
    message("Mail saved to '", o365$get_drafts()$properties$displayName, "' folder of account ", o365$properties$mail, ".")
    mail_lst
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
       md_body = FALSE,
       signature = FALSE,
       automated_notice = FALSE,
       ...)
}

#' @rdname mail
#' @export
mail_dataset <- function(df, font.size = 11, ..., plain = FALSE) {
  if (!is.data.frame(df)) {
    warning("mail_dataset() can only handle data.frames", call. = FALSE)
    return("")
  }
  if (plain == TRUE) {
    df %>%
      tbl_flextable(font.size = 10,
                    font.family = "Courier New",
                    autofit.fullpage = FALSE,
                    logicals = c("TRUE", "FALSE"),
                    format.dates = "yyyy-mm-dd",
                    format.NL = FALSE,
                    column.names.bold = FALSE,
                    row.names.bold = FALSE) %>%
      htmltools_value(class = "")
  } else {
    df %>%
      tbl_flextable(font.size = font.size, ..., autofit.fullpage = FALSE) %>%
      htmltools_value()
  }
}

#' @rdname mail
#' @importFrom blastula add_image
#' @param image_path path of image
#' @param width required width of image, must be in CSS style such as "200px"
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
    warning("Image file does not exist: ", normalizePath(path), call. = FALSE)
    return("")
  }
}

#' @rdname mail
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
                           background = colourpicker("certeroze2"),
                           md_body = TRUE,
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
#' @param path save location for attachment(s)
#' @param search an ODATA filter, ignores `sort`
#' @param folder email folder to search in
#' @param n maximum number of email to list
#' @param sort initial sorting
#' @param overwrite logical to indicate whether existing local files should be overwritten
#' @export
download_mail_attachment <- function(path = getwd(),
                                     search = NULL,
                                     folder = "Postvak IN",
                                     n = 50,
                                     sort = "received desc",
                                     overwrite = TRUE) {
  o365 <- get_outlook365(error_on_fail = TRUE)
  folder <- o365$get_folder(folder)
  if (!is.null(search)) {
    message("'search' provided, ignoring 'sort = ", sort, "'")
  }
  mails <- folder$list_emails(n = n, by = sort, search = search)
  has_attachment <- sapply(mails,
                           function(m) {
                             m$properties$hasAttachments &&
                               sapply(m$list_attachments(), function(a) a$properties$name != "0")
                           })
  if (!any(has_attachment)) {
    warning("No mails with attachment found in the ", n, " mails searched (sort = \"", sort, "\")", call. = FALSE)
    return(NULL)
  }
  mails <- mails[which(has_attachment)]
  mails_txt <- sapply(mails,
                      function(m) {
                        p <- m$properties
                        dt <- as.POSIXct(gsub("T", " ", p$receivedDateTime), tz = "UTC")
                        paste0(crayon::bold(format2(dt, "ddd d mmm yyyy HH:MM")),
                               crayon::bold("u"),
                               " ", p$subject, " ",
                               crayon::blue(p$from$emailAddress$address),
                               "\n",
                               paste0("    - ", sapply(m$list_attachments(), function(a) paste0(a$properties$name, " (", round(a$properties$size / 1024) , " kB)")),
                                      collapse = "\n")
                        )
                      })
  if (interactive()) {
    # pick mail
    mail_int <- menu(mails_txt,
                     graphics = interactive(),
                     title = paste0(length(mails), "/", n,
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
      att_int <- menu(sapply(att, function(a) paste0(a$properties$name, " (", round(a$properties$size / 1024), " kB)")),
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
