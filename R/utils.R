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

globalVariables(c(".", "att", "end", "is_all_day", "start"))

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
  # inherits() returns FALSE for NULL, so no need to check is.null(account)
  inherits(account, "ms_object") &&
    tryCatch(!is.null(account$create_email), error = function(e) FALSE) &&
    is.function(account$create_email)
}


#' @importFrom certestyle format2
format_filename <- function(mail, attachment, filename, add_seq_if_exists = FALSE) {
  if (is.null(filename)) {
    return(attachment$properties$name)
  }

  dt <- as.POSIXct(format(as.POSIXct(gsub("T", " ", mail$properties$receivedDateTime),
                                     tz = "UTC"),
                          tz = "Europe/Amsterdam"))# replace text items with relative properties
  filename <- gsub("{date}", format2(dt, "yyyymmdd"), filename, fixed = TRUE)
  filename <- gsub("{time}", format2(dt, "HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{datetime}", format2(dt, "yyyymmdd_HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{date_time}", format2(dt, "yyyymmdd_HHMMSS"), filename, fixed = TRUE)
  filename <- gsub("{name}", mail$properties$from$emailAddress$name, filename, fixed = TRUE)
  filename <- gsub("{address}", mail$properties$from$emailAddress$address, filename, fixed = TRUE)
  filename <- gsub("{original}", attachment$properties$name, filename, fixed = TRUE)

  # replace invalid filename characters
  filename <- trimws(gsub("[\\/:*?\"<>|]+", "_", filename))
  original_filename <- filename

  # add extension if missing
  ext <- gsub(".*([.].*)$", "\\1", attachment$properties$name)
  if (filename %unlike% paste0(ext, "$")) {
    filename <- paste0(filename, ext)
  }

  # add sequence number if file exists
  if (add_seq_if_exists == TRUE) {
    i <- 1
    while (file.exists(filename)) {
      filename <- paste0(original_filename, "_", i)
      i <- i + 1
      # add extension if missing
      if (filename %unlike% paste0(ext, "$")) {
        filename <- paste0(filename, ext)
      }
    }
  }

  filename
}

only_valid_attachments <- function(attachments, search, skip_inline) {
  if (is.null(search) && skip_inline == FALSE) {
    return(attachments)
  }
  is_valid <- vapply(FUN.VALUE = logical(1),
                     attachments,
                     function(a, s = search)
                       ifelse(!is.null(search),
                              a$properties$name %like% s,
                              TRUE) &&
                       ifelse(skip_inline == TRUE,
                              !a$properties$isInline,
                              TRUE))
  attachments[which(is_valid)]
}

size_formatted <- function(size) {
  if (size > 1024 ^ 2) {
    size <- paste0(round(size / 1024 ^ 2, 1), " MB")
  } else if (size > 1024) {
    size <- paste0(round(size / 1024, 0), " kB")
  } else {
    size <- paste0(size, " B")
  }
  size
}
