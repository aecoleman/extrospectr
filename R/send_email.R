#' Send Email
#'
#' Uses COM dispatch to communicate with MS Outlook and send an email.
#'
#' @param to character, vector of email addresses which will be receiving the email. Each element must contain the \code{@} symbol, with one or more characters on either side.
#' @param cc character (optional), vector of email addresses which will be carbon-copied (CCed) on the email. If provided, each element must contain the \code{@} symbol, with one or more characters on either side.
#' @param subject character, the subject line for the email. If this has length > 1, an error will be returned.
#' @param body character, the body of the email. If a vector of length > 1 is provided, the elements will be treated as lines of the email, and will be combined using \code{paste0(..., collapse = "\r\n")}.
#' @param body_html logical, does the argument to body contain HTML code that should control how the email is rendered? Default is \code{FALSE}.
#' @param attachments character (optional), paths to files that are to be attached. If any of these files does not exist, an error will be returned.
#' @param display_only logical, should the email be displayed only (and not sent)? Default is \code{FALSE}, which sends the email without displaying it first.
#'
#' @return NULL (invisibly)
#' @export
#'
send_email <- function(to, cc = NULL, subject, body, body_html = FALSE, attachments = NULL, display_only = FALSE) {

  # Start Guards

  # All elements of `to` must be email addresses
  stopifnot(all(.is_email_address(to)))

  # All elements of `cc` must be email addresses
  # Note: this is vacuously true if cc has no elements
  stopifnot(is.null(cc) || all(.is_email_address(cc)))

  # The argument `subject` must be a character vector of length 1
  stopifnot(length(subject) == 1L && identical(typeof(subject), 'character'))

  # The argument `body` must be a character vector
  stopifnot(identical(typeof(body), 'character'))

  # All elements of `attachments` must be paths to files that exist
  stopifnot(is.null(attachments) || all(file.exists(attachments)))

  # If body has more than 1 element, treat each element as a line of the email.
  if (length(body) > 1L) {
    body <- paste0(body, collapse = '\r\n')
  }

  # Initialize the COM connection to Outlook
  ol <- RDCOMClient::COMCreate('Outlook.Application')

  # Create a new mail item
  new_mail <- ol$CreateItem(0)

  # Add `To` for email
  new_mail[['To']] <-
    .make_outlook_address_list(to)

  # Add `CC` for email, if it was provided
  if (!is.null(cc)) {
    new_mail[['CC']] <-
      .make_outlook_address_list(cc)
  }

  # Add the email subject
  new_mail[['subject']] <- subject

  # TODO: Create (or locate) function to check whether the body contains one or
  # more HTML tags
  if (body_html == TRUE) {
    new_mail[['HTMLbody']] <- body
  } else {
    new_mail[['body']] <- body
  }

  # Add attachments, if any were provided
  if (!is.null(attachments)) {
    # We are passing the file path to Outlook, which does not know anything
    # about the working directory of our R session, so we must create an
    # absolute path that Outlook can use to locate the file we intend to add
    attachments <-  normalizePath(attachments, mustWork = TRUE)

    # Add each attachment, one by one
    for (attachment in attachments) {
      new_mail[['Attachments']]$Add(attachment)
    }

    rm(attachment)

  }

  # Check if the user wants to send the email, or only to display it
  if (display_only == TRUE) {
    new_mail$Display()
  } else {
    new_mail$Send()
  }

  return(invisible(NULL))

}

#' Make Outlook Address List
#'
#' @param address character, vector of email addresses
#'
#' @return character singleton
#'
.make_outlook_address_list <- function(address) {

  # TODO: Decide whether this is a good idea...
  address <- unlist(stringr::str_split(address, pattern = ';'))

  stopifnot(all(.is_email_address(address)))

  return(paste0(address, collapse = ';'))

}

#' Is Email Address
#'
#' @param address character, vector of email addresses
#'
#' @return logical vector of same length as address
#'
#' @examples
#'
#' .is_email_address("test@example.com")
#'
#' .is_email_address("test[at]example[dot]com")
#'
.is_email_address <- function(address) {

  # Check to make sure that each address includes exactly one @ symbol
  has_one_at <- stringr::str_count(address, '@') == 1L

  # Check to make sure that each address includes one or more characters before
  # the @ symbol
  has_local_part <- stringr::str_detect(address, '^[^@]+@')

  # Check to make sure that each address includes one or more characters after
  # the @ symbol
  has_domain_part <- stringr::str_detect(address, '@[^@]+$')

  return(has_one_at & has_local_part & has_domain_part)

}