# Add comment
# 
# Author: Andrie
#----------------------------------------------------------------------------------


testFile <- "braid.R"
resultFile <- "braidPPT.ppt"


context("braid a ppt file")

test_that("Creating braidppt file is possible", {
      
      if(file.exists(testFile)) file.remove(testFile)
      if(file.exists(resultFile)) file.remove(resultFile)
      
      filePath <- getwd()
      b <- as.braid(path=filePath, fileInner=testFile)
      
      braidpptNew(b)
      braidpptNewSlide(b, title="This is the title", text="And this is the text")
      braidpptInsertImage(b, file="sinewave.png")
      braidpptSave(b, resultFile)
      braidpptClose(b)
      
      braidSave(b)
      
      expect_true(file.exists(testFile))
      
      test <- scan(file=testFile, what="character", sep="\n")
      rest <- c(
          "ppt <- pptNew()", 
          "ppt <- pptNewSlide(ppt, title=\"This is the title\", text=\"And this is the text\")",
          "ppt <- pptInsertImage(ppt, file=\"sinewave.png\")",
          "PPT.SaveAs(ppt, \"braidPPT.ppt\")",
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
      
    })

test_that("Compiling braidppt creates a ppt file", {
      
      expect_false(file.exists(resultFile))
      
      braidCompilePPT(fileOuter=testFile)
      
      expect_true(file.exists(resultFile))
      
    })


