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
pkg_env$o365 <- NULL

globalVariables(c(".", "att"))

validate_mail_address <- function(x) {
  x <- trimws(x)
  x <- x[!is.na(x) & x != ""]
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

is_valid_o365 <- function(account) {
  !is.null(account) &&
    tryCatch(!is.null(account$create_email), error = function(e) FALSE) &&
    is.function(account$create_email) &&
    inherits(account, "ms_object")
}


#' @importFrom certestyle format2
format_filename <- function(mail, attachment, filename) {
  if (is.null(filename)) {
    return(attachment$properties$name)
  }

  dt <- as.POSIXct(gsub("T", " ", mail$properties$receivedDateTime), tz = "UTC")
  # replace text items with relative properties
  filename <- gsub("{date}", format2(dt, "yyyymmdd"), filename, fixed = TRUE)
  filename <- gsub("{time}", format2(dt, "HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{datetime}", format2(dt, "yyyymmdd_HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{date_time}", format2(dt, "yyyymmdd_HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{name}", mail$properties$from$emailAddress$name, filename, fixed = TRUE)
  filename <- gsub("{address}", mail$properties$from$emailAddress$address, filename, fixed = TRUE)
  filename <- gsub("{original}", attachment$properties$name, filename, fixed = TRUE)

  # replace invalid filename characters
  filename <- trimws(gsub("[\\/:*?\"<>|]+", "_", filename))

  # add extension if missing
  ext <- gsub(".*([.].*)$", "\\1", attachment$properties$name)
  if (filename %unlike% paste0(ext, "$")) {
    filename <- paste0(filename, ext)
  }

  filename
}

only_valid_attachments <- function(attachments, search) {
  if (is.null(search)) {
    return(attachments)
  }
  is_valid <- vapply(FUN.VALUE = logical(1),
                     attachments,
                     function(a, s = search) a$properties$name %like% s)
  attachments[which(is_valid)]
}
