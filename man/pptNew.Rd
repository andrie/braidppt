\name{pptNew}
\alias{pptNew}
\title{Create new PowerPoint slide and applies template.}
\usage{
  pptNew(template = "PentaLibra.potx", method =
  "RDCOMClient")
}
\arguments{
  \item{template}{Character string: File location and
  template name.  If not a full file path, searches for the
  file in the default path for Microsoft templates.  See
  \code{\link{msTemplatePath}} for details}

  \item{method}{Character string indicating connection
  method to COM.  Currently supported options are "rcom"
  and "RDCOMClient"}
}
\value{
  A pointer to a ppt object
}
\description{
  Create new PowerPoint slide and applies template.
}
\examples{
if(isCOMsupported(method="RDCOMClient")){
ppt <- pptNew()
}
}
\seealso{
  Other PowerPoint: \code{\link{pptClose}},
  \code{\link{pptNewSlide}}, \code{\link{pptSave}}
}

