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

my_secrets_file <- tempfile(fileext = ".yaml")
Sys.setenv(secrets_file = my_secrets_file)
writeLines(c("mail.auto_cc: ''",
             "mail.auto_bcc: ''"),
           my_secrets_file)

test_that("mail works", {
  expect_message(mail("test body", "test subject", to = "to@domain.com",
                      cc = NULL, bcc = NULL, reply_to = NULL,
                      account = NULL))
   m <- mail(mtcars[1:5, ], "test subject", to = "to@domain.com",
             cc = NULL, bcc = NULL, reply_to = NULL,
             account = NULL, send = FALSE)
  expect_s3_class(m, "certe_mail")
  expect_s3_class(m, "blastula_message")
  expect_output(print(m))
})

test_that("download works", {
  expect_message(download_mail_attachment(account = NULL))
})
