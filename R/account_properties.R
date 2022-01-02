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
#' This function creates a connection to Outlook Business 365 and saves this connection to the `certemail` package environment. The `get_*()` functions retrieve properties from this connection.
#' @param tenant the tenant to use for [Microsoft365R::get_business_outlook()]
#' @param error_on_fail a [logical] to indicate whether an error must be thrown if no connection can be made
#' @param account a Microsoft 365 account to use for looking up properties. This has to be an object as returned by [connect_outlook365()] or [Microsoft365R::get_business_outlook()].
#' @rdname account_properties
#' @name account_properties
#' @export
#' @importFrom Microsoft365R get_business_outlook
connect_outlook365 <- function(tenant = read_secret("mail.tenant"), error_on_fail = FALSE) {
  if (tenant == "") {
    tenant <- NULL
  }
  if (is.null(pkg_env$o365)) {
    # not yet connected to Microsoft 365, so set it up
    tryCatch({
      if (is.null(tenant)) {
        pkg_env$o365 <- suppressWarnings(suppressMessages(get_business_outlook()))
      } else {
        pkg_env$o365 <- suppressWarnings(suppressMessages(get_business_outlook(tenant = tenant)))
      }
      message("Connected to Microsoft 365 as ", get_name_and_mail_address(), ".")
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
  if (inherits(account, "R6")) {
    account$properties$displayName
  } else {
    NA_character_
  }
}

#' @rdname account_properties
#' @export
get_name_and_job_title <- function(account = connect_outlook365()) {
  if (inherits(account, "R6")) {
    paste(account$properties$displayName, "|", account$properties$jobTitle)
  } else {
    NA_character_
  }
}

#' @rdname account_properties
#' @export
get_name_and_mail_address <- function(account = connect_outlook365()) {
  if (inherits(account, "R6")) {
    paste0(account$properties$displayName, " (", account$properties$mail, ")")
  } else {
    NA_character_
  }
}

#' @rdname account_properties
#' @export
get_mail_address <- function(account = connect_outlook365()) {
  if (inherits(account, "R6")) {
    account$properties$mail
  } else {
    NA_character_
  }
}

#' @rdname account_properties
#' @export
get_inbox_name <- function(account = connect_outlook365()) {
  if (inherits(account, "R6")) {
    account$get_inbox()$properties$displayName
  } else {
    NA_character_
  }
}
