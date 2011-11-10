\name{tabIndentLevels}
\alias{tabIndentLevels}
\title{Determines indentation level of text, based on tab characters (\\t) at start of line.}
\usage{
  tabIndentLevels(text, sep = "\r")
}
\arguments{
  \item{text}{Character string, with paragraph breaks
  indicated by \\r or \\n}

  \item{sep}{Character Separator to indicated paragraph
  breaks.  Defaults to \\r, the PowerPoint standard}
}
\value{
  a list consisting of two elements: text and levels
}
\description{
  Splits the input text at paragraph boundaries (determined
  by \\t or \\n), counts the number of leading tabs in each
  paragraph, removes the tab stops, and returns text with
  \\r as collapsed string.
}
\keyword{internal}

