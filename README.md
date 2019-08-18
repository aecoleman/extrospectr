# extrospectr

<!-- badges: start -->
<!-- badges: end -->

The goal of extrospectr is to give R users the ability to retrieve and dispatch 
emails using Microsoft Outlook. The name is pseudo-latin, `extro-` meaning 
(roughly) "outwards" and `spect` meaning "to observe" (or, one might say "to 
look"). It's a bad name, and I apologize.

Note that this package requires Microsoft Windows, and requires that 
you have Microsoft Outlook installed. For many situations, it may be better to 
use [gmailr](http://gmailr.r-lib.org/). However, in a situation where you must 
send or receive an email from your organization email, this may be a good 
solution.

Many thanks are due to [Duncan Lang](https://github.com/omegahat/), whose 
[RDCOMClient](https://github.com/omegahat/RDCOMClient) package made this package 
possible. In many ways, this is simply a wrapper for `RDCOMClient` so that the 
R user can abstract away the more difficult aspects working with the DCOM API.

## Installation

You can install the development version of `extrospectr` from github with:

``` r
# install.packages("remotes")

remotes::install_github("omegahat/RDCOMClient")
remotes::install_github("aecoleman/extrospectr")
```

## Reading Outlook Inbox

To read the contents of an inbox or other mail folder, you can use the following:

``` r
# Attach extrospectr and RDCOMClient libraries
library(extrospectr)

read_inbox(folder_path = "name@example.com/Inbox")

```

If you have multiple addresses linked to Outlook, or have a nested folder structure, you can use the `folder_path` argument to specify which folder of which email address to read. For example: `read_inbox(folder_path = "you@gmail.com/[gmail]/Starred")`.

Which will return a `data.frame` with eight columns. An example (transformed to a `tibble`) is shown below:

``` r

 #  EntryID  Subject  CreationTime     LastModificatio~ MessageClass Sender ReceivedTime     Content 
 #   <chr>    <chr>    <dttm>           <dttm>           <chr>        <chr>  <dttm>           <chr>   
 # 1 EF00000~ Write a~ 2019-08-13 21:36 2019-08-11 15:47 IPM.Note     autom~ 2019-08-11 15:47 "Your f~
 # 2 EF00000~ Thank y~ 2019-08-13 21:36 2019-08-11 18:23 IPM.Note     navih~ 2019-08-11 18:23 "Dear A~
 # 3 EF00000~ Thank y~ 2019-08-13 21:36 2019-08-11 18:24 IPM.Note     navih~ 2019-08-11 18:24 "Dear A~
 # 4 EF00000~ You hav~ 2019-08-13 21:36 2019-08-11 18:31 IPM.Note     team@~ 2019-08-11 18:31 " <http~
 # 5 EF00000~ Thank y~ 2019-08-13 21:36 2019-08-11 18:33 IPM.Note     norep~ 2019-08-11 18:33 " <http~

```
Many of the columns are adequately explained by their names, but a few deserve 
some special explanation. Note that the character columns are truncated.

The `EntryID` column stores a unique identifier that can be used to specify a particular email to other functions. 

The `CreationTime` column seems to store the time when Outlook downloaded the email from the server.

The `Content` column stores the first 255 characters of the body of the email. 

It is worth commenting that objects other than Outlook `MailItem`s can end up in your inbox. For example, a calendar invite, or a notice that a sent message was returned undeliverable. The function filters to only `MailItems`, and so these objects won't be displayed.

## Saving Attachments

Description to come...

## Notes

One potential "gotcha" is that `RDCOMClient` must be attached for the functions 
provided in this package to work properly. If you attach `extrospectr`, then 
`RCDOMClient` will be attached as well. However, if you attempt to call 
`extrospectr::read_inbox(...)` while `RDCOMClient` is not attached, you will 
receive an error.

```
> extrospectr::read_inbox('me@example.com/Inbox')
Error in createCOMReference(<pointer: 0x000002363aeacb08>, "COMIDispatch") : 
  could not find function "createCOMReference"
Called from: getCOMInstance(name, force = TRUE, silent = TRUE)
```