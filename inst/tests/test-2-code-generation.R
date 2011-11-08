# Add comment
# 
# Author: Andrie
#----------------------------------------------------------------------------------


context("Code generation helper functions")

test_that(".q() quotes correctly", {
      
      expect_equal(.q("abc"), "\"abc\"")
      expect_equal(.q(123), 123)
      expect_equal(.q(letters[1:3]), "c(\"a\", \"b\", \"c\")")
      expect_equal(.q(1:3),  "c(1, 2, 3)")
    })

test_that(".qarg() and .cqarg() returns correct value",{
      
      expect_true(is.null(.qarg("file"))) 
      expect_equal(.qarg("file", "braid.ppt"), "file=\"braid.ppt\"")
      expect_true(is.null(.cqarg("file")))
      expect_equal(.cqarg("file", "braid.ppt"), ", file=\"braid.ppt\"")
      
      expect_equal(.qarg("size", c(1, 2, 3, 4)), "size=c(1, 2, 3, 4)")
      
      expect_equal(
          paste("ppt <- pptInit(", .qarg("template", NULL), ")", sep=""),
          "ppt <- pptInit()"
      )
      
      expect_equal(
          paste("ppt <- pptInit(", .qarg("template", "std.pot"), ")", sep=""),
          "ppt <- pptInit(template=\"std.pot\")"
      )
      
      expect_equal(
          paste("ppt <- pptInit(", .qarg("size", c(1, 2, 3)), ")", sep=""),
          "ppt <- pptInit(size=c(1, 2, 3))"
      )
      
      expect_equal(
          paste("ppt <- pptInit(", .qarg("s", c(1, 2, 3)), .cqarg("s2", c(4, 5, 6)),")", sep=""),
          "ppt <- pptInit(s=c(1, 2, 3), s2=c(4, 5, 6))"
      )

    })  

