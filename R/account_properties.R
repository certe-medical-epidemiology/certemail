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
#' @param email email address of the user, or a shared mailbox
#' @param force logical to indicate whether the token must be retrieved again. This is useful for changing accounts; calling [outlook_connect()] again after an initial call with a different email address will allow all `certemail` functions to use the newly set account (since the default value for `force` is `TRUE`, while all `certemail` functions call [outlook_connect()] with `force = FALSE`).
#' @param ... arguments passed on to [`get_microsoft365_token()`][certeprojects::get_microsoft365_token()]
#' @param account the Microsoft 365 Outlook account object, e.g. as returned by [outlook_connect()]
#' @details The [get_certe_name()] and [get_certe_name_and_job_title()] functions first look for the [secret][certetoolbox::read_secret()] `user.{user}.fullname` and falls back to [get_name()] if it is not available.
#' @rdname account_properties
#' @name account_properties
#' @export
#' @importFrom AzureGraph create_graph_login
outlook_connect <- function(email = read_secret("mail.auto_from"),
                            force = TRUE,
                            ...) {
  if (!"certeprojects" %in% rownames(utils::installed.packages())) {
    stop("This requires the 'certeprojects' package")
  }
  if (is.null(pkg_env$m365_getmail) || isTRUE(force)) {
    # not yet connected to Outlook in Microsoft 365, so set it up
    pkg_env$m365_getmail <- create_graph_login(token = certeprojects::get_microsoft365_token(scope = "outlook", ...))$
      get_user(email = email)$
      get_outlook()
    message("Connected to Microsoft 365 Outlook as ",
            get_name(account = pkg_env$m365_getmail),
            " (", get_mail_address(account = pkg_env$m365_getmail), ").")
  }
  # this will auto-renew authorisation when due
  return(invisible(pkg_env$m365_getmail))
}

#' @rdname account_properties
#' @export
get_name <- function(account = outlook_connect(force = FALSE)) {
  get_property(account, "displayName")
}

#' @rdname account_properties
#' @export
get_department <- function(account = outlook_connect(force = FALSE)) {
  get_property(account, "department")
}

#' @rdname account_properties
#' @export
get_address <- function(account = outlook_connect(force = FALSE)) {
  out <- get_property(account, c("streetAddress", "postalCode", "city"))
  if (any(!is.na(out))) {
    out <- paste(get_property(account, c("streetAddress", "postalCode", "city")), collapse = " ")
  }
  out
}

#' @rdname account_properties
#' @export
get_phone_numbers <- function(account = outlook_connect(force = FALSE)) {
  get_property(account, c("mobilePhone", "businessPhones"))
}

#' @rdname account_properties
#' @export
get_mail_address <- function(account = outlook_connect(force = FALSE)) {
  get_property(account, "mail")
}

#' @rdname account_properties
#' @export
get_inbox_name <- function(account = outlook_connect(force = FALSE)) {
  get_property(account, "displayName", "get_inbox")
}

#' @rdname account_properties
#' @export
get_drafts_name <- function(account = outlook_connect(force = FALSE)) {
  get_property(account, "displayName", "get_drafts")
}

#' @rdname account_properties
#' @param user Certe user ID, to look up the [secret][certetoolbox::read_secret()] `user.{user}.fullname`
#' @export
get_certe_name <- function(user = Sys.info()["user"],
                           account = outlook_connect(force = FALSE)) {
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
                          account = outlook_connect(force = FALSE)) {
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
                                         account = outlook_connect(force = FALSE)) {
  fullname <- get_certe_name(user = user, account = account)
  jobtitle <- get_job_title(account = account)
  if (!all(is.na(fullname)) && !all(is.na(jobtitle))) {
    paste(fullname, "|", jobtitle)
  } else {
    NA_character_
  }
}

#' @rdname account_properties
#' @export
get_certe_location <- function(account = outlook_connect(force = FALSE)) {
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
get_certe_signature <- function(account = outlook_connect(force = FALSE), plain = FALSE) {
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
                  get_certe_name_and_job_title(account = account), "  \n",
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
