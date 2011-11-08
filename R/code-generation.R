# Add comment
# 
# Author: Andrie
#----------------------------------------------------------------------------------


# Helper functions to correctly quote arguments 

# Embeds argument (of length 1) in quotes if class is character
.qone <- function(x){
  if(is.character(x)) sprintf("\"%s\"", x) else x
}

# Embeds argument in quotes if class is character
# Vector arguments are collapsed with paste
.q <- function(x){
  if (length(x)==1) .qone(x) else sprintf("c(%s)", paste(.qone(x), collapse=", "))
}


# Quote argument - if value is NULL, returns NULL: this makes for simpler code
.qarg <- function(arg, value){
  if(missing(value) || is.null(value)) NULL else paste(arg, .q(value), sep="=")
}

# Same as .qarg, but
.cqarg <- function(arg, value){
  x <- .qarg(arg, value)
  if(is.null(x)) NULL else sprintf(", %s", x)
}

