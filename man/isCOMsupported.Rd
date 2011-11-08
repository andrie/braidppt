\name{isCOMsupported}
\alias{isCOMsupported}
\title{Tests whether R architecture is 32 bit Windows.}
\usage{
  isCOMsupported(ppt = NULL, method = ppt$method)
}
\arguments{
  \item{ppt}{A pointer to a ppt object}

  \item{method}{Character string indicating connection
  method to COM.  Currently supported options are "rcom"
  and "RDCOMClient"}
}
\value{
  TRUE if architecture is \code{i386} else FALSE
}
\description{
  rcom is only available on 32 but windows.
}
\examples{
isCOMsupported(method="rcom")
isCOMsupported(method="RDCOMClient")
}
\keyword{Internal}

