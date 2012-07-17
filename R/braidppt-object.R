# Braid object definition
# 
# Author: Andrie
#------------------------------------------------------------------------------


#' Creates object of class braidppt. This is an extension of class braid, specifically to create PowerPoint output.
#' 
#' A braid object describes the data and meta-data in the survey that will be analysed by the analysis and reporting functions in the braid package.
#' 
#' @param path Character vector: File path where latex output will be saved
#' @param graphics Character vector: File path relative to \code{pathOutput} where to save graphics (this gets appended to pathOutput.)
#' @param fileOuter Filename of latex outline file (used by \code{\link{braidCompile}}
#' @param fileInner Filename where latex output will be saved (This gets appended to \code{pathOutput})
#' @param counterStart The starting number for a counter used to store graphs, defaults to 1
#' @param defaultPlotSize Numeric vector of length 2: Plot size in inches, e.g. c(4, 3)
#' @param dpi Dots per inch, passed to ggsave()
#' @param outputType Character string specifying the destination of output: "latex", "ppt" or "device".  If "device", graphs are sent to the default device (typically the RGgui plot terminal)
#' @param graphicFormat Device type for saving graphic plots.  Currently only pdf and wmf is supported.
#' @return A list object of class braid
#' @export 
as.braidppt <- function(
    path = tempdir(),
    graphics = "graphics",
    fileOuter = "braidppt.R",
    fileInner = fileOuter,
    counterStart = 1,
    defaultPlotSize = c(5,3),
    dpi = 600,
    outputType = "ppt",
    graphicFormat = "wmf"
    
){
  if (!file_test("-d", path)) {
    stop(paste("The file path does not exist:", path))
  }
  x <- as.braid(
      path=path,
      graphics=graphics,
      fileOuter=fileOuter,
      fileInner=fileInner,
      counterStart=counterStart,
      defaultPlotSize=defaultPlotSize,
      dpi=dpi,
      outputType=outputType,
      graphicFormat=graphicFormat
  )
  class(x) = c("braidppt", "braid")
  x
}

#' Tests that object is of class braid.
#' 
#' @param x Object to be tested
#' @export
is.braidppt <- function(x){
  inherits(x, "braidppt")
}


