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

#' Connect to Outlook Business 365
#'
#' These functions create a connection to Outlook Business 365 and saves the connection to the `certemail` package environment. The `get_*()` functions retrieve properties from this connection.
#' @param tenant the tenant to use for [Microsoft365R::get_business_outlook()]
#' @param app_id the Azure app id to use for [Microsoft365R::get_business_outlook()]
#' @param error_on_fail a [logical] to indicate whether an error must be thrown if no connection can be made
#' @param account a Microsoft 365 account to use for looking up properties. This has to be an object as returned by [connect_outlook365()] or [Microsoft365R::get_business_outlook()].
#' @details The [get_certe_name()] and [get_certe_name_and_job_title()] functions first look for the [secret][certetoolbox::read_secret()] `user.{user}.fullname` and falls back to [get_name()] if it is not available.
#' @rdname account_properties
#' @name account_properties
#' @export
#' @importFrom Microsoft365R get_business_outlook
connect_outlook365 <- function(shared_mbox_email = read_secret("mail.auto_from"),
                               tenant = read_secret("mail.tenant"),
                               app_id = read_secret("mail.app_id"),
                               auth_type = read_secret("mail.auth_type"),
                               error_on_fail = FALSE) {
  # see here: https://docs.microsoft.com/en-us/graph/permissions-reference
  scopes <- c("Mail.ReadWrite", "Mail.Send", "User.Read")
  if (shared_mbox_email == "") {
    shared_mbox_email <- NULL
  } else {
    scopes <- c(scopes, "Mail.Send.Shared", "Mail.ReadWrite.Shared")
  }
  if (tenant == "") {
    tenant <- NULL
  }
  if (app_id == "") {
    app_id <- get(".microsoft365r_app_id", envir = asNamespace("Microsoft365R"))
  } else {
    scopes <- c("Mail.ReadWrite", "Mail.Send", "User.Read",
                "Mail.Send.Shared", "Mail.ReadWrite.Shared")
  }
  if (auth_type == "") {
    auth_type <- NULL
  }
  if (is.null(pkg_env$o365)) {
    # not yet connected to Microsoft 365, so set it up
    tryCatch({
      if (is.null(tenant)) {
        pkg_env$o365 <- suppressWarnings(suppressMessages(get_business_outlook(shared_mbox_email = shared_mbox_email,
                                                                               scopes = scopes,
                                                                               app = app_id,
                                                                               auth_type = auth_type)))
      } else {
        pkg_env$o365 <- suppressWarnings(suppressMessages(get_business_outlook(tenant = tenant,
                                                                               shared_mbox_email = shared_mbox_email,
                                                                               scopes = scopes,
                                                                               app = app_id,
                                                                               auth_type = auth_type)))
      }
      message("Connected to Microsoft 365 as ",
              get_name(account = pkg_env$o365),
              " (", get_mail_address(account = pkg_env$o365), ").")
    }, warning = function(w) {
      return(invisible())
    }, error = function(e, fail = error_on_fail) {
      if (isTRUE(fail)) {
        stop("Could not connect to Microsoft 365: ", paste0(e$message, collapse = ", "), call. = FALSE)
      } else {
        warning("Could not connect to Microsoft 365: ", paste0(e$message, collapse = ", "), call. = FALSE)
      }
      return(NULL)
    })
  }
  if (isTRUE(error_on_fail) && is.null(pkg_env$o365)) {
    stop("Could not connect to Microsoft 365.", call. = FALSE)
  }
  # this will auto-renew authorisation when due
  return(invisible(pkg_env$o365))
}

#' @rdname account_properties
#' @export
get_name <- function(account = connect_outlook365()) {
  get_property(account, "displayName")
}

#' @rdname account_properties
#' @export
get_department <- function(account = connect_outlook365()) {
  get_property(account, "department")
}

#' @rdname account_properties
#' @export
get_address <- function(account = connect_outlook365()) {
  out <- get_property(account, c("streetAddress", "postalCode", "city"))
  if (any(!is.na(out))) {
    out <- paste(get_property(account, c("streetAddress", "postalCode", "city")), collapse = " ")
  }
  out
}

#' @rdname account_properties
#' @export
get_phone_numbers <- function(account = connect_outlook365()) {
  get_property(account, c("mobilePhone", "businessPhones"))
}

#' @rdname account_properties
#' @export
get_mail_address <- function(account = connect_outlook365()) {
  get_property(account, "mail")
}

#' @rdname account_properties
#' @export
get_inbox_name <- function(account = connect_outlook365()) {
  get_property(account, "displayName", "get_inbox")
}

#' @rdname account_properties
#' @export
get_drafts_name <- function(account = connect_outlook365()) {
  get_property(account, "displayName", "get_drafts")
}

#' @rdname account_properties
#' @param user Certe user ID, to look up the [secret][certetoolbox::read_secret()] `user.{user}.fullname`
#' @export
get_certe_name <- function(user = Sys.info()["user"],
                           account = connect_outlook365()) {
  if (!is.null(user)) {
    secr <- suppressWarnings(read_secret(paste0("user.", user, ".fullname")))
    if (secr != "") {
      return(secr)
    }
  }
  get_name(account = account)
}


#' @rdname account_properties
#' @export
get_job_title <- function(user = Sys.info()["user"],
                          account = connect_outlook365()) {
  if (!is.null(user)) {
    secr <- suppressWarnings(read_secret(paste0("user.", user, ".jobtitle")))
    if (secr != "") {
      return(secr)
    }
  }
  get_property(account, "jobTitle")
}

#' @rdname account_properties
#' @export
get_certe_name_and_job_title <- function(user = Sys.info()["user"],
                                         account = connect_outlook365()) {
  fullname <- get_certe_name(user = user, account = account)
  jobtitle <- get_job_title(account = account)
  if (!is.na(fullname) & !is.na(jobtitle)) {
    paste(fullname, "|", jobtitle)
  } else {
    NA_character_
  }
}

#' @rdname account_properties
#' @export
get_certe_location <- function(account = connect_outlook365()) {
  out <- get_property(account, "streetAddress")[1]
  if (!is.na(out)) {
    out <- trimws(paste("Locatie", gsub("[^a-zA-Z ]", "", out)))
  }
  out
}

#' @rdname account_properties
#' @param plain a [logical] to indicate whether textual formatting should not be applied
#' @importFrom certestyle colourpicker
#' @export
get_certe_signature <- function(account = connect_outlook365(), plain = FALSE) {
  if (!is_valid_o365(account)) {
    return(NULL)
  }
  phones <- get_phone_numbers(account = account)
  if (all(is.na(phones))) {
    phones <- character(0)
  } else {
    phones <- paste(phones, collapse = " | ")
  }

  location <- get_certe_location(account = account)
  if (all(is.na(location))) {
    location <- character(0)
  }

  if (isTRUE(plain)) { # no markdown
    out <- paste0(c("Met vriendelijke groet,\n\n",
                    get_certe_name_and_job_title(account = account),
                    get_mail_address(account = account),
                    "\nCERTE",
                    # "Afdeling ", get_department(account = account)[1L],
                    "Postbus 909 | 9700 AX Groningen | certe.nl",
                    phones,
                    location),
                  collapse = "  \n")
  } else {
    out <- paste0('<div style="font-family: Calibri, Verdana !important; margin-top: 0px !important;">',
                  "Met vriendelijke groet,\n\n",
                  '<div class="white"></div>\n\n',
                  get_certe_name_and_job_title(account = account),"  \n",
                  "[", get_mail_address(account = account), "](mailto:", get_mail_address(account = account), ")\n\n",
                  # logo:
                  '<div style="font-family: \'Arial Black\', \'Calibri\', \'Verdana\' !important; font-weight: bold !important; color: ',
                  colourpicker("certeblauw"),
                  ' !important; font-size: 16px !important;" class="certelogo">CERTE</div>',
                  # "Afdeling ", get_department(account = account)[1L], "<br>",
                  'Postbus 909 | 9700 AX Groningen | <a href="https://www.certe.nl">certe.nl</a><br>',
                  phones, "<br>",
                  location,
                  "</div>")
  }
  structure(out, class = c("certe_signature", "character"))
}

#' @noRd
#' @importFrom htmltools html_print HTML
#' @export
print.certe_signature <- function(x, ...) {
  html_print(HTML(commonmark::markdown_html(paste(as.character(x), collapse = "\n"), ...)))
}

get_property <- function(account, property_names, account_fn = NULL) {
  if (is_valid_o365(account)) {
    if (!is.null(account_fn)) {
      account <- account[[account_fn]]()
    }
    out <- unique(as.character(unname(unlist(account$properties[c(property_names)]))))
    if (length(out) == 0 || all(is.na(out))) {
      out <- NA_character_
    }
  } else {
    out <- NA_character_
  }
  return(out)
}
