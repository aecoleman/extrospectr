% Generated by roxygen2: do not edit by hand
% Please edit documentation in R/send_email.R
\name{.is_email_address}
\alias{.is_email_address}
\title{Is Email Address}
\usage{
.is_email_address(address)
}
\arguments{
\item{address}{character, vector of email addresses}
}
\value{
logical vector of same length as address
}
\description{
Is Email Address
}
\examples{

.is_email_address("test@example.com")

.is_email_address("test[at]example[dot]com")

}
