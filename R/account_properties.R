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

#' Work with Outlook 365
#'
#' After connecting using [connect_outlook()], the `get_*()` functions retrieve properties from this connection.
#' @param account the Microsoft 365 Outlook account object, e.g. as returned by [connect_outlook()]
#' @details The [get_certe_name()] and [get_certe_name_and_job_title()] functions first look for the [secret][certetoolbox::read_secret()] `user.{user}.fullname` and falls back to [get_name()] if it is not available.
#' @rdname account_properties
#' @name account_properties
#' @export
get_name <- function(account = connect_outlook()) {
  get_property(account, "displayName")
}

#' @rdname account_properties
#' @export
get_department <- function(account = connect_outlook()) {
  get_property(account, "department")
}

#' @rdname account_properties
#' @export
get_address <- function(account = connect_outlook()) {
  out <- get_property(account, c("streetAddress", "postalCode", "city"))
  if (any(!is.na(out))) {
    out <- paste(get_property(account, c("streetAddress", "postalCode", "city")), collapse = " ")
  }
  out
}

#' @rdname account_properties
#' @export
get_phone_numbers <- function(account = connect_outlook()) {
  get_property(account, c("mobilePhone", "businessPhones"))
}

#' @rdname account_properties
#' @export
get_mail_address <- function(account = connect_outlook()) {
  get_property(account, "mail")
}

#' @rdname account_properties
#' @export
get_inbox_name <- function(account = connect_outlook()) {
  get_property(account, "displayName", "get_inbox")
}

#' @rdname account_properties
#' @export
get_drafts_name <- function(account = connect_outlook()) {
  get_property(account, "displayName", "get_drafts")
}

#' @rdname account_properties
#' @param user Certe user ID, to look up the [secret][certetoolbox::read_secret()] `user.{user}.fullname`
#' @export
get_certe_name <- function(user = Sys.info()["user"],
                           account = connect_outlook()) {
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
                          account = connect_outlook()) {
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
                                         account = connect_outlook()) {
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
get_certe_location <- function(account = connect_outlook()) {
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
get_certe_signature <- function(account = connect_outlook(), plain = FALSE) {
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
                    "\nCERTE",
                    paste0("Afdeling ", trimws(gsub("Certe", "", read_secret("department.name")))),
                    read_secret("department.mail"),
                    read_secret("department.phone"),
                    read_secret("department.address")),
                  collapse = "  \n")
  } else {
    out <- paste0('<div style="font-family: Calibri, Verdana !important; margin-top: 0px !important;">',
                  "Met vriendelijke groet,<br>",
                  '<div class="white"></div><br>',
                  get_certe_name_and_job_title(account = account), "  <br><br>",
                  # logo:
                  '<div style="font-family: \'Arial Black\', \'Calibri\', \'Verdana\' !important; font-weight: bold !important; color: ', colourpicker("certeblauw"), ' !important; font-size: 16px !important;" class="certelogo">CERTE</div>',
                  # rest:
                  paste0("Afdeling ", trimws(gsub("Certe", "", read_secret("department.name")))), "<br>",
                  read_secret("department.mail.html"), "<br>",
                  read_secret("department.phone"), "<br>",
                  read_secret("department.address.html"), "<br>",
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

#' @importFrom certeprojects connect_outlook
#' @export
certeprojects::connect_outlook
