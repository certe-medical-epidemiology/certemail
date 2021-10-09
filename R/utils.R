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

#' @importFrom Microsoft365R get_business_outlook
get_outlook365 <- function(error_on_fail = FALSE) {
  if (is.null(pkg_env$o365)) {
    # not yet connected to Microsoft 365, so try this
    tryCatch(
      pkg_env$o365 <<- suppressMessages(get_business_outlook(tenant = read_secret("tenant"))),
      error = function(e, fail = error_on_fail) {
        if (fail == TRUE) {
          stop("Could not connect to Microsoft 365: ", e$message, call. = FALSE)
        } else {
          warning("Not connected to Microsoft 365", call. = FALSE)
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


read_secret <- function(property) {
  file <- Sys.getenv("secrets_file")
  if (file == "" || !file.exists(file)) {
    stop("System environmental variable 'secrets_file' not set")
  }
  file_lines <- lapply(strsplit(readLines(Sys.getenv("secrets_file")), ":"), trimws)
  # set name to list item
  names(file_lines) <- sapply(file_lines, function(l) l[1])
  # strip name from vector
  file_lines <- lapply(file_lines, function(l) l[c(2:length(l))])
  # get property
  out <- file_lines[[property]]
  if (is.null(out)) {
    return(NULL)
  }
  # split on comma
  trimws(strsplit(out, ",")[[1]])
}

colourpicker <- function(...) {
  # to-do: get from certestyle package
  "white"
}

`%like%` <- function(x, pattern) {
  # to-do: get from certetools package
  grepl(x = x, pattern = pattern, ignore.case = TRUE)
}

concat <- function(...) {
  # to-do: get from certetools package
  paste0(..., collapse = "")
}
