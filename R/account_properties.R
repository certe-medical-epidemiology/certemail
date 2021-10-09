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

#' Retrieve specific account properties
#'
#' @rdname account_properties
#' @export
get_name <- function() {
  o365 <- get_outlook365()
  o365$properties$displayName
}

#' @rdname account_properties
#' @export
get_name_and_job_title <- function() {
  o365 <- get_outlook365()
  paste(o365$properties$displayName, "|", o365$properties$jobTitle)
}

#' @rdname account_properties
#' @export
get_mail_address <- function() {
  o365 <- get_outlook365()
  o365$properties$mail
}
