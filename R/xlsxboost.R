# Provides functions to ease manipulations of .xlsx files in R

#' Package initialization function
#' 
#' Initializes package by preventing usage of awt on mac; to be called before use
#' of any functions herein.
#' 
#' @param initializeRJava Set true if you are experiencing problems with initialization
#' of RJava (win32/64); attempts will be made to rectify installation. Default setting is false.
#' @export
#' @examples
#' 
#' xlsx.initialize() # we may now use functionality exposed in package 'xlsxboost'
#' 
xlsx.initialize <- function(initializeRJava = FALSE) {
  OS = .Platform$OS.type
  if(OS == "unix") {
    if(Sys.info()["sysname"] == "Linux")
      OS = "linux"
    else
      OS = "mac"
  }
  if(OS == "mac")
    Sys.setenv(NOAWT = 1) # prevents usage of awt - required on Mac
  
  # Let's now initialize RJava
  if(initializeRJava)
    xlsx.initializeRJava()
}

#' RJava initialization function
#' 
#' Initializes RJava on a system by resetting JAVA_HOME
#' 
#' @export
#' @examples
#' 
#' xlsx.initializeRJava() # initializes JAVA_HOME to "" for proper installation
#' 
xlsx.initializeRJava <- function() {
  if (Sys.getenv("JAVA_HOME") != "")
    Sys.setenv(JAVA_HOME = "")
  library(rJava)
}

#' Opens a .xlsx file provided a relative filepath from getwd()
#' 
#' Platform-independent method to open a .xlsx file; requires a relative filepath from getwd()
#' as well as capability of opening the document from terminal (Excel / LibreOffice / etc.)
#' 
#' @param filename the name of the file to open
#' 
#' @export
#' @examples
#' 
#' require('xlsx')
#' wb <- createWorkbook(type="xlsx")
#' sheet <- xlsx::createSheet(wb, sheetName = "example")
#' saveWorkbook(wb, "test.xlsx")
#' xlsx.openFile("test.xlsx")
#' 
xlsx.openFile<-function(filename = NULL)
{
  absolute.path = paste(getwd(), "/", filename, sep="")
  
  if(.Platform$OS.type=="windows"){
    shell.exec(absolute.path)
  }
  else if(.Platform$OS.type=="unix"){
    system(paste("open ", absolute.path, sep=""))
  }    
}

#' Adds a header to a .xlsx file
#' 
#' Makes use of numerous parameters including the sheet, text, HTML heading level, and color
#' to display a header on an excel workbook
#' 
#' @param wb The excel workbook
#' @param sheet The excel sheet (created by xlsxBoost)
#' @param value The header value specified
#' @param level The HTML header level
#' @param color The hex color code
#' @param startRow The row to start the header on
#' @param startCol The column to start the header on
#' 
#' @export
#' @examples
#' 
#' wb <- createWorkbook(type="xlsx")
#' sheet <- xlsx::createSheet(wb, sheetName = "example")
#' xlsx.addHeader(wb, sheet, value="Header 1",level=1, color="black")
#' saveWorkbook(wb, "example.xlsx")
#' xlsx.openFile("example.xlsx") # view the file
#' 
xlsx.addHeader<-function(wb, sheet, value="Header", level=1, color="#FFFFFF",
                         startRow=NULL, startCol=2)
{
  if(color=="black")
    color="white" # black and white are inverted in package 'xlsx'
  
  H1 <- CellStyle(wb) + 
    Font(wb,  heightInPoints=22, color=color, isBold=TRUE, underline=0)
  H2 <- CellStyle(wb) + 
    Font(wb, heightInPoints=18, color=color, isItalic=FALSE, isBold=TRUE, underline=0)
  H3 <- CellStyle(wb) + 
    Font(wb, heightInPoints=16, color=color, isItalic=TRUE, isBold=TRUE, underline=0)
  H4 <- CellStyle(wb) + 
    Font(wb, heightInPoints=16, color=color, isItalic=TRUE, isBold=FALSE, underline=0)
  H5 <- CellStyle(wb) + 
    Font(wb, heightInPoints=14, color=color, isItalic=TRUE, isBold=FALSE, underline=0)
  H6 <- CellStyle(wb) + 
    Font(wb, heightInPoints=12, color=color, isItalic=TRUE, isBold=FALSE, underline=0)
  
  # Now, append the row to the sheet
  if(is.null(startRow)){
    rows <- getRows(sheet)
    startRow = length(rows) + 1  
  } 
  
  # Create the sheet title and subtitle
  rows <- createRow(sheet,rowIndex=startRow)
  sheetTitle <- createCell(rows, colIndex=startCol)
  setCellValue(sheetTitle[[1,1]], value)
  
  if(level==1)      xlsx::setCellStyle(sheetTitle[[1,1]], H1)
  else if(level==2) xlsx::setCellStyle(sheetTitle[[1,1]], H2)
  else if(level==3) xlsx::setCellStyle(sheetTitle[[1,1]], H3)
  else if(level==4) xlsx::setCellStyle(sheetTitle[[1,1]], H4)
  else if(level==5) xlsx::setCellStyle(sheetTitle[[1,1]], H5)
  else if(level==6) xlsx::setCellStyle(sheetTitle[[1,1]], H6)  
}


#' Adds a plot to a .xlsx file
#' 
#' Makes use of numerous parameters including the sheet, text, start row/column, and 
#' width/height in order to display a plot in an excel workbook provided a plot function.
#' 
#' @param wb The excel workbook
#' @param sheet The excel sheet (created by xlsxboost)
#' @param plotFunction The function (i.e. ggplot2(...)) to show in Excel
#' @param startRow The row to start the header on
#' @param startCol The column to start the header on
#' @param width The width of the plot
#' @param height The height of the plot
#' 
#' @export
#' @examples
#' 
#' wb <- createWorkbook(type="xlsx")
#' sheet <- xlsx::createSheet(wb, sheetName = "example")
#' x = seq(from = -10*pi, to = 10*pi, by = 0.1)
#' pf <- function() {plot(sin(x))}
#' xlsx.addPlot(wb, sheet, pf)
#' saveWorkbook(wb, "example2.xlsx")
#' xlsx.openFile("example2.xlsx") # view the file
#' 
xlsx.addPlot<-function(wb, sheet, plotFunction, startRow=NULL, startCol=2,
                       width=480, height=480, ...)
  
{
  # Step 1: plot and create the PNG
  png(filename = "plot.png", width = width, height = height,...)
  plotFunction()
  dev.off() 
  
  # Append plot to sheet if startRow is not specified
  if(is.null(startRow)){
    rows <- getRows(sheet)
    startRow = length(rows) + 1
  } 
  
  # Add the file created previously
  addPicture("plot.png", sheet=sheet,  startRow = startRow, startColumn = startCol) 
  
  # Status
  res <- file.remove("plot.png")
  res
}