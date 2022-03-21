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
#' @param error_on_fail a [logical] to indicate whether an error must be thrown if no connection can be made
#' @param account a Microsoft 365 account to use for looking up properties. This has to be an object as returned by [connect_outlook365()] or [Microsoft365R::get_business_outlook()].
#' @details The [get_full_name()] and [get_full_name_and_job_title()] functions also take into account the [secrets][certetoolbox::read_secret()] set as `mail.{user}.prefix_name` and `mail.{user}.suffix_name`.
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
get_job_title <- function(account = connect_outlook365()) {
  get_property(account, "jobTitle")
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
#' @param prefix text to place before the name
#' @param suffix text to place after the name
#' @export
get_full_name <- function(account = connect_outlook365(),
                          prefix = read_secret(paste0("mail.", Sys.info()["user"], ".prefix_name")),
                          suffix = read_secret(paste0("mail.", Sys.info()["user"], ".suffix_name"))) {
  name <- get_name(account = account)
  if (!is.na(name)) {
    name <- trimws(gsub(" +", " ", paste(prefix, name, suffix)))
    name <- gsub(" ([^a-zA-Z0-9])", "\\1", name)
  }
  name
}

#' @rdname account_properties
#' @export
get_full_name_and_job_title <- function(account = connect_outlook365(),
                                        prefix = read_secret(paste0("mail.", Sys.info()["user"], ".prefix_name")),
                                        suffix = read_secret(paste0("mail.", Sys.info()["user"], ".suffix_name"))) {
  fullname <- get_full_name(account = account, prefix = prefix, suffix = suffix)
  jobtitle <- get_job_title(account = account)
  if (!is.na(fullname) & !is.na(jobtitle)) {
    paste(fullname, "|", jobtitle)
  } else {
    NA_character_
  }
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
