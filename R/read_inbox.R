
#' Read Inbox
#'
#' @param folder_path character
#' @param with_attachments_only logical, should the results be filtered to include only emails with attachments?
#' @param count_attachments logical, should the number of attachments be returned?
#' @param table_filter character (optional), additional filters to be applied
#' @param whole_body logical, should the entire body of the email be returned? Default is FALSE
#'
#' @return data.frame
#' @export
#'
read_inbox <- function(folder_path, with_attachments_only = FALSE, count_attachments = FALSE, table_filter = NULL, whole_body = FALSE) {

  ol <- RDCOMClient::COMCreate('Outlook.Application')$GetNamespace('MAPI')

  vec.folder_path <- unlist(strsplit(folder_path, split = '/'))

  # Store initial pointer for the root of the folder tree
  ptr.folder <- ol

  # Traverse the folder tree until we arrive at the desired folder
  for (folder in vec.folder_path) {
    # Replace the pointer for the current folder with the pointer for the folder
    # we want to go to next
    ptr.folder <- ptr.folder$folders(folder)

  }

  # Gets pointer to Table representation of folder contents, filtering to only
  # mailItems (i.e., no meeting requests or notifications, or any other
  # non-email item that may end up in the inbox
  ptr.fldr_tbl <-
    ptr.folder$GetTable(
      '@SQL="http://schemas.microsoft.com/mapi/proptag/0x001a001e" = \'IPM.Note\'')

  if (with_attachments_only == TRUE) {

    ptr.fldr_tbl <-
      ptr.fldr_tbl$Restrict('@SQL= "urn:schemas:httpmail:hasattachment" = 1')

  }

  if (!is.null(table_filter)) {

    ptr.fldr_tbl <-
      ptr.fldr_tbl$Restrict(table_filter)

  }

  n <- ptr.fldr_tbl$GetRowCount()

  # Add Columns to Table

  ptr.fldr_tbl$Columns()$Add('SenderEmailAddress')

  ptr.fldr_tbl$Columns()$Add('ReceivedTime')

  ptr.fldr_tbl$Columns()$Add('Content')

  ptr.fldr_tbl$MoveToStart()

  tbl_ary <-
    ptr.fldr_tbl$GetArray(n)

  out <-
    tbl_ary %>%
    .format_table_array()

  out

}

#' Count Attachments
#'
#' @param EntryID entryID referring to an Outlook object
#' @param ol pointer to an active Outlook MAPI namespace
#'
#' @return integer
#'
.count_attachments <- function(EntryID, ol = NULL) {

  if (is.null(ol)) {
    ol <- RDCOMClient::COMCreate('Outlook.Application')$GetNamespace('MAPI')
  }

  if (length(EntryID) > 1L) {

    out <-
      EntryID %>%
      purrr::map_int(
        .count_attachments,
        ol)

  } else {
    out <-
      ol$GetItemFromID(EntryID)$Attachments()$Count()
  }

  out

}

#' Read Body
#'
#' @inheritParams .count_attachments
#'
#' @return character
#' @export
#'
.read_body <- function(EntryID, ol = NULL) {

  if (is.null(ol)) {
    ol <- RDCOMClient::COMCreate('Outlook.Application')$GetNamespace('MAPI')
  }

  if (length(EntryID) > 1L) {

    out <-
      EntryID %>%
      purrr::map_int(
        .read_body,
        ol)

  } else {
    out <-
      ol$GetItemFromID(EntryID)$Body()
  }

  out

}

#' Lookup Exchange Sender
#'
#' @inheritParams .count_attachments
#'
#' @return character
#'
.lookup_exchange_sender <- function(EntryID, ol = NULL) {

  if (is.null(ol)) {
    ol <- RDCOMClient::COMCreate('Outlook.Application')$GetNamespace('MAPI')
  }

  if (length(EntryID) > 1L) {

    out <-
      EntryID %>%
      purrr::map_int(
        .lookup_exchange_sender,
        ol)

  } else {

    ptr.sender <- ol$GetItemFromID(EntryID)$Sender()

    user_type <-
      extrospectr::address_type %>%
      dplyr::filter(
        id == ptr.sender$AddressEntryUserType())

    out <-
      switch(
        user_type,
        olExchangeDistributionListAddressEntry = ptr.sender$GetExchangeDistributionList()$PrimarySMTPAddress(),
        olExchangeUserAddressEntry = ptr.sender$GetExchangeUser()$PrimarySMTPAddress(),
        )
  }

  out

}