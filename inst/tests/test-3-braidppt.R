# Add comment
# 
# Author: Andrie
#----------------------------------------------------------------------------------


filePath <- tempdir()
testFile <- "braid.R"
resultFile <- "braidPPT.ppt"


context("braidppt")

test_that("Creating and compiling a braidppt is possible", {
      
      file.exists(filePath)
      
      if(file.exists(file.path(filePath, testFile))) file.remove(file.path(filePath, testFile))
      if(file.exists(file.path(filePath, resultFile))) file.remove(file.path(filePath, resultFile))
      
      b <- as.braidppt(path=filePath, fileOuter=testFile)
      expect_is(b, "braidppt")
      expect_is(b, "braid")

      png(file.path(filePath, "sinewave.png"))
      curve(sin, -pi, pi)
      dev.off()
      on.exit({
            file.remove(file.path(filePath, "sinewave.png"))
            #PPT.Close(ppt)
          })
      
      braidpptNew(b)
      braidpptNewSlide(b, title="This is the title", text="And this is the text")
      
      braidpptInsertImage(b, file=file.path(filePath, "sinewave.png"))
      braidpptSave(b, resultFile)
      braidpptClose(b)
      
      braidSave(b)
      
      expect_true(file.exists(file.path(filePath, testFile)))
      
      test <- scan(file=file.path(filePath, testFile), what="character", sep="\n")
      rest <- c(
          "ppt <- pptNew()", 
          "ppt <- pptNewSlide(ppt, title=\"This is the title\", text=\"And this is the text\")",
          sprintf("ppt <- pptInsertImage(ppt, file=\"%s\")", encodeString(file.path(filePath, "sinewave.png"))),
          sprintf("PPT.SaveAs(ppt, \"%s\")", encodeString(normalizePath(file.path(filePath, resultFile), mustWork=FALSE))),
          "PPT.Close(ppt)"
      )
      
      if(!identical(test, rest)){
        cat("\n\nResults of scan:\n\n")
        cat(paste(test, collapse="\n"))
        cat("\n\nExpected results:\n\n")
        cat(paste(rest, collapse="\n"))
        cat("\n\n")
      }
      
      expect_equal(test, rest)
      
      expect_false(file.exists(file.path(filePath, resultFile)))
      braidCompile(b)#, fileOuter=testFile)
      expect_true(file.exists(file.path(filePath, resultFile)))
      
      
      
    })

