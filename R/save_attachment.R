#' Save All Attachments
#'
#' Save all attachments of a particular email, given some identifier for that
#' email
#'
#' @inheritParams save_attachment
#' @param save_to_dir character, path to the local directory where the files are to be saved
#'
#' @return character, vector of paths where attachments have been saved
#' @export
#'
save_all_attachments <- function(mailItem, save_to_dir, must_work = FALSE) {

  stopifnot(dir.exists(save_to_dir))

  # Check whether the user supplied a pointer already. If not, use what the user
  # supplied to get a pointer
  if (.is_MailItem(mailItem) == TRUE) {
    ptr.MailItem <- mailItem
  } else {
    ns.mapi <- RDCOMClient::COMCreate('Outlook.Application')$GetNamespace('MAPI')
    ptr.MailItem <- ns.mapi$GetItemFromID(mailItem)
  }

  # Get number of attachments to save
  n <- ptr.MailItem$Attachments()$Count()

  # Initialize output vector
  out <- vector(length = n, mode = 'character')

  for (i in seq_len(n)) {

    out[i] <-
      .save_attachment(
          ptr.MailItem = ptr.MailItem,
          attachment = i,
          save_path = save_to_dir,
          must_work = must_work)

  }

  return(out)

}

#' Save Attachment
#'
#' Save a particular attachment to a location, given an identifier for the Email
#' the desired file is attached to, and some identifier for which attached file
#' to download.
#'
#' @param mailItem either a COMIDispatch pointer to a MailItem object, or an EntryID for a MailItem
#' @param attachment the name (character) or index (integer) of the attachment to be saved
#' @param save_path character, path to the location where the file should be saved
#'
#' @return character, the path where the file was saved, or \code{NA_character_}
#' if a check determines that the file does not exist, meaning that the file was
#' not saved
#' @export
#'
save_attachment <- function(mailItem, attachment, save_path, must_work = FALSE) {

  # Check whether the use supplied a pointer, or supplied an EntryID
  if (.is_MailItem(x) == TRUE) {

    ptr.MailItem <- mailItem

  } else if (is.character(x)) {
    ns.mapi <- RDCOMClient::COMCreate('Outlook.Application')$GetNamespace('MAPI')

    ptr.MailItem <- ns.mapi$GetItemByID(x)
  }

  # Argument `attachment` must be a character, or must be a positive integer
  stopifnot(is.character(attachment) || is.integer(attachment) & all(attachment > 0))

  attachment_name <- ptr.MailItem$Attachments(attachment)$DisplayName()

  # If the user provided a directory which exists, then we will save the file in
  # that directory, otherwise we will save it
  if (dir.exists(save_path)) {
    save_path <- file.path(save_path, attachment_name)
  }

  ptr.MailItem$Attachments(attachment)$SaveAsFile(save_path)

  ret_val <- ifelse(file.exists(save_path), save_path, NA_character_)

  stopifnot(must_work == FALSE || !is.na(ret_val))

  # Return the file path if the attachment was saved successfully, NA otherwise
  return(ret_val)

}

#' Is MailItem
#'
#' Returns \code{TRUE} if x is a pointer to a MailItem object, and \code{FALSE}
#' otherwise.
#'
#' @param x object(s) to be checked
#'
#' @return logical of length equal to the length of \code{x}
#'
.is_MailItem <- function(x) {

  if (length(x) > 1) {

    purrr::map_lgl(
      x,
      .is_MailItem)

  } else {

    'COMIDispatch' == class(x) && x$Class() == 43

  }
}