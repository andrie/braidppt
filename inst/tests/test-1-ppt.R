# Add comment
# 
# Author: Andrie
#----------------------------------------------------------------------------------

context("ppt helper functions")

test_that("msTemplatePath returns correct file path",{
      
      userPath <- path.expand("~/")
      testPath <- file.path(userPath, "AppData", "Roaming", "Microsoft", "Templates")
      testFile <- list.files(testPath, pattern="^[[:alnum:]].*.ot.*$")[1]
      expFile <- list.files(testPath, pattern="^[[:alnum:]].*.ot.*$", full.names=TRUE)[1]
      expFile <- normalizePath(expFile)
      
      expect_equal(msTemplatePath(testFile), expFile)
      expect_equal(msTemplatePath(expFile), expFile)
      expect_equal(msTemplatePath(testFile, path=testPath), expFile)

})



context("Initialise ppt and add text slides")

test_that("pptNew creates new ppt using rcom", {
      
      if(.Platform$r_arch=="i386"){
        ppt <- pptNew(method="rcom")
        expect_is(ppt, "list")
        expect_equal(ppt$method, "rcom")
        expect_is(ppt$ppt, "COMObject")
        expect_is(ppt$pres, "COMObject")
        
        ppt <- pptClose(ppt)
        rm(ppt)
      }
#      print(dput(ppt))
      
    })

test_that("pptNew creates new ppt using RDCOMClient", {
      
      ppt <- pptNew()
      expect_is(ppt, "list")
      expect_equal(ppt$method, "RDCOMClient")
      expect_is(ppt$ppt, "COMIDispatch")
      expect_is(ppt$pres, "COMIDispatch")
      
      ppt <- pptClose(ppt)
      
#      print(dput(ppt))
      
    })


test_that("newSlide adds slides", {
      
      ppt <- pptNew()

      ppt <- pptNewSlide(ppt) # This should be blank
      ppt <- pptNewSlide(ppt, "Title", subtitle="Subtitle")
      ppt <- pptNewSlide(ppt, "Title only")
      ppt <- pptNewSlide(ppt, "Title with text", "Some text\rSome more text\rEven more text")

      expect_is(ppt, "list")
      expect_equal(ppt$method, "RDCOMClient")
      expect_is(ppt$ppt, "COMIDispatch")
      expect_is(ppt$pres, "COMIDispatch")
      
      pptClose(ppt)
    
    })
      
context("Save ppt")

test_that("ppt is created and saved",{
      
      testFile <- "testRppt.ppt"
      
      if(file.exists(testFile)) file.remove(testFile)
      
      ppt <- pptNew()
      ppt <- pptNewSlide(ppt) # This should be blank
      ppt <- pptNewSlide(ppt, "Title", subtitle="Subtitle")
      ppt <- pptNewSlide(ppt, "Title only slide")
      ppt <- pptNewSlide(ppt, "Title with text", "Some text\rSome more text\rEven more text")
      ppt <- pptSave(ppt, testFile)
      ppt <- pptClose(ppt)
      
      expect_true(file.exists(testFile))
      
    })
