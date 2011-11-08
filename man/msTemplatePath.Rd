\name{msTemplatePath}
\alias{msTemplatePath}
\title{Guess where template file is found in user path.}
\usage{
  msTemplatePath(file, path = NULL)
}
\arguments{
  \item{file}{Character string: file name (with or without
  full path)}

  \item{path}{Character string: path}
}
\description{
  If \code{file} is a valid file, then does nothing.
  Otherwise attempts to find the file in the Microsoft
  template folder - typically somehting like
  \code{C:\\Users\\...\\AppData\\Roaming\\Microsoft\\Templates\\}
}
\keyword{internal}

