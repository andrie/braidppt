\name{braidpptPlot}
\alias{braidpptPlot}
\title{Inserts image on slide.}
\usage{
  braidpptPlot(braid, plotcode,
    filename = braidFilename(braid),
    width = braid$defaultPlotSize[1],
    height = braid$defaultPlotSize[2], Qid = NA)
}
\arguments{
  \item{plotcode}{A plot object (either ggplot or lattice)}

  \item{width}{Numeric: width of plot in inches}

  \item{height}{Numeric: height of plot in inches}

  \item{Qid}{Question ID}

  \item{braid}{Braid object}

  \item{filename}{Filename of image to attach}
}
\description{
  Inserts image on slide.
}
\seealso{
  Other braidPPT: \code{\link{braidpptClose}},
  \code{\link{braidpptInsertImage}},
  \code{\link{braidpptNew}},
  \code{\link{braidpptNewSlide}},
  \code{\link{braidpptSave}}
}

