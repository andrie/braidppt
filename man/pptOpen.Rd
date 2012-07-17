\name{pptOpen}
\alias{pptOpen}
\title{Opens existing PowerPoint presentations.}
\usage{
  pptOpen(file, method = c("RDCOMClient", "rcom"))
}
\arguments{
  \item{file}{Character string: Name of a file}

  \item{method}{Character string indicating connection
  method to COM.  Currently supported options are "rcom"
  and "RDCOMClient"}
}
\value{
  A pointer to a ppt object
}
\description{
  Opens existing PowerPoint presentations.
}
\seealso{
  Other PowerPoint: \code{\link{pptClose}},
  \code{\link{pptNew}}, \code{\link{pptNewSlide}},
  \code{\link{pptSave}}
}

