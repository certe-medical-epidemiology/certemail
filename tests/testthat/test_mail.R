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
             "mail.auto_bcc: ''",
             "mail.export_path: ''"),
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

  m2 <- mail(mail_image(image_path = system.file("test.jpg", package = "certemail")),
             "test subject", to = "to@domain.com",
             cc = NULL, bcc = NULL, reply_to = NULL,
             account = NULL, send = FALSE, markdown = FALSE)
  expect_s3_class(m2, "certe_mail")

  expect_true(mail_on_error(1 + 1 == 2, account = NULL))
  expect_message(suppressWarnings(mail_on_error(1 + 1 == nonexistingobject, account = NULL)))
})

test_that("download works", {
  expect_message(download_mail_attachment(account = NULL))
})

test_that("properties return NA", {
  expect_identical(get_name(account = NULL), NA_character_)
  expect_identical(get_name_and_job_title(account = NULL), NA_character_)
  expect_identical(get_name_and_mail_address(account = NULL), NA_character_)
  expect_identical(get_mail_address(account = NULL), NA_character_)
  expect_identical(get_department(account = NULL), NA_character_)
  expect_identical(get_inbox_name(account = NULL), NA_character_)
  expect_identical(get_drafts_name(account = NULL), NA_character_)
})

test_that("utils work", {
  expect_error(validate_mail_address("nothing"))
  expect_error(validate_mail_address("nothing@test"))
  expect_error(validate_mail_address("test.com"))
  expect_identical(validate_mail_address("a@test.com"), "a@test.com")
  expect_null(validate_mail_address(NULL))
  expect_null(validate_mail_address(""))

  expect_false(is_valid_o365(NULL))
  expect_false(is_valid_o365("test"))
})
