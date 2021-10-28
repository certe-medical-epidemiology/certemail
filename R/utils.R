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

# this is the package environment. The Microsoft 365 connection will be saved to this env.
pkg_env <- new.env(hash = FALSE)

globalVariables(c("."))

#' @importFrom Microsoft365R get_business_outlook
#' @importFrom certetoolbox read_secret
get_outlook365 <- function(error_on_fail = FALSE) {
  if (is.null(pkg_env$o365)) {
    # not yet connected to Microsoft 365, so try this
    tryCatch({
      tenant <- read_secret("tenant")
      if (tenant == "") {
        tenant <- NULL
      }
      pkg_env$o365 <<- suppressWarnings(suppressMessages(get_business_outlook(tenant = tenant)))
      message("Connected to Microsoft 365 as ", get_name_and_mail_address(), ".")
    }, warning = function(w) {
      return(invisible())
    }, error = function(e, fail = error_on_fail) {
      if (fail == TRUE) {
        stop("Could not connect to Microsoft 365: ", paste0(e$message, collapse = ", "), call. = FALSE)
      } else {
        warning("Could not connect to Microsoft 365: ", paste0(e$message, collapse = ", "), call. = FALSE)
      }
      return(NULL)
    })
  }
  # this will auto-renew authorisation when due
  return(pkg_env$o365)
}

validate_mail_address <- function(x) {
  x <- trimws(x)
  x <- x[!is.na(x)]
  if (length(x) == 0) {
    return(NULL)
  }
  for (i in seq_len(length(x))) {
    if (!grepl("^[a-z0-9._-]+@[a-z0-9._-]+[.][a-z]+$", x[i], ignore.case = TRUE)) {
      stop("Invalid email address: '", x[i], "'")
    }
  }
  x
}
