# Add comment
# 
# Author: Andrie
#----------------------------------------------------------------------------------



#' Initialises braidppt.
#' 
#' @param braid Braid object
#' @param template Character string: location and file name of template to apply
#' @family braidPPT
#' @export 
braidpptNew <- function(braid, template=NULL){
  text <- paste0("ppt <- pptNew(", .qarg("template", template), ")")
  braid::braidWrite(braid, text)
  invisible(NULL)
}

#' Adds new slide.
#' 
#' @inheritParams braidpptNew
#' @param title Slide title
#' @param text Slide text
#' @param subtitle Slide subtitle
#' @param filename Filename of image to attach
#' @param size Size of image
#' @family braidPPT
#' @export 
braidpptNewSlide <- function(braid, title=NULL, text=NULL, subtitle=NULL, filename=NULL, 
    size=NULL){
  text <- paste0("ppt <- pptNewSlide(ppt", .cqarg("title", title), .cqarg("text", text),
      .cqarg("subtitle", subtitle), .cqarg("file", filename), .cqarg("size", size), ")" )
  braid::braidWrite(braid, text)
  invisible(NULL)
}

#' Inserts image on slide.
#' 
#' @inheritParams braidpptNewSlide
#' @family braidPPT
#' @export 
braidpptInsertImage <- function(braid, filename, size=NULL){
  #filename <- gsub("/", "\\\\", as.character(filename))
  #path     <- gsub("/", "\\\\", as.character(path))
  text <- paste0("ppt <- pptInsertImage(ppt", 
      .cqarg("file", encodeString(filename)), 
      .cqarg("size", size), 
      #.cqarg("path", encodeString(path)), 
      ")"
  )
  braid::braidWrite(braid, text)
  invisible(NULL)
}




#' Inserts image on slide.
#' 
#' @inheritParams braidpptNewSlide
#' @param plotcode A plot object (either ggplot or lattice)
#' @param width Numeric: width of plot in inches
#' @param height Numeric: height of plot in inches
#' @param Qid Question ID
#' @family braidPPT
#' @export 
braidpptPlot <- function(braid, plotcode, filename=braidFilename(braid), 
    width=braid$defaultPlotSize[1], 
    height=braid$defaultPlotSize[2], Qid=NA){
  
  fullFilename <- file.path(braid$pathGraphics, filename)
  if(grepl("pdf$", fullFilename)) stop("Unable to insert pdf images into PowerPoint")
  
  braidpptInsertImage(braid, filename=fullFilename, size=NULL)
  braid:::braidAppendPlot(braid, plotcode, filename, width, height, Qid)
  invisible(NULL)
}


#' Saves ppt.
#' 
#' @inheritParams braidpptNew
#' @param filename Character string: Name of PowerPoint file to save.
#' @param path file path to filename
#' @family braidPPT
#' @export 
braidpptSave <- function(braid, filename, path=braid$pathLatex){
  f <- normalizePath(file.path(path, basename(filename)), mustWork=FALSE)
  ef <- encodeString(f)
  text <- sprintf("PPT.SaveAs(ppt, \"%s\")", ef)
  #browser()
  #message(text)
  braid::braidWrite(braid, text)
  invisible(NULL)
}

#' Closes ppt.
#' 
#' @inheritParams braidpptNew
#' @family braidPPT
#' @export 
braidpptClose <- function(braid){
  text <- "PPT.Close(ppt)"
  braid::braidWrite(braid, text)
  invisible(NULL)
}


#' Compiles braid object to ppt.
#' 
#' Sources the code created with \code{braidppt}.
#' 
#' @param b A braid object
#' @param fileOuter Location of R script file
#' @param ... not used
#' @method braidCompile braidppt
#' @export
braidCompile.braidppt <- function(b, fileOuter=b$fileInner, ...){
  wd <- getwd()
  on.exit(setwd(wd))
  setwd(b$pathLatex)
  source(fileOuter)
}

#' Compiles braid object to ppt.
#' 
#' Sources the code created with \code{braidppt}.
#' 
#' @inheritParams braidCompile.braidppt
#' @export
braidCompilePPT <- function(b, fileOuter=b$fileInner, ...){
  .Deprecated("braidCompile.braidppt")
  braidCompile(b, fileOuter=b$fileInner, ...)
}

