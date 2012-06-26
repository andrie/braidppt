\name{pptInsertImage}
\alias{pptInsertImage}
\title{Add picture (image file) to slide.}
\usage{
  pptInsertImage(ppt, file = NULL,
    size = c(0.1, 0.1, 0.9, 0.9))
}
\arguments{
  \item{ppt}{A ppt object, see \code{\link{pptNew}}}

  \item{file}{Location and file name of a plot}

  \item{size}{A numeric vector of size 4, indicating c(x1,
  y1, x2, y2) expressed as percentage of page size}
}
\description{
  The default position is in the middle of the slide, at
  the size of the original image.
}
\examples{
##if(isRcomSupported()){
##  ppt <- newSlide(ppt, "Slide with graphic", "This is really interesting",
##    plot=plotIs("company_language.wmf"), size=c(0.5, 0.25, 1, 0.75))
##}
}

