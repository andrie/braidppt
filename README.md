# braidppt

braidppt is an extension of the `braid` package in R.  It allows report writing from R to PowerPoint.

It uses (and imports) the `R2PPT` package by Wayne Jones.

braidppt provides the following wrapper functions for manipulating ppt objects:

* pptNew
* pptClose
* pptSave
* pptNewSlide
* pptInsertImage

It extends `braid` by providing functions for:

* braidpptNew
* braidpptNewSlide
* braidCompile

---

This package depends on an interface to windows COM. You can do this using either of two packages:

* rcom
* RDCOMClient

I have found that RDCOMClient works on windows 64-bit. The package documentation is at:

http://www.omegahat.org/RDCOMClient/

To install, use:

install.packages("RDCOMClient", repos = "http://www.omegahat.org/R")

You may have to install from source to get it to work:

install.packages("RDCOMClient", repos = "http://www.omegahat.org/R", type="source")
