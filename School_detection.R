School_detection <- function(){

#################################################################################
# This script opens ed04_reviewed.EV, creates an editable line at 5m range, 
# and uses this line as the exclude below line threshold in the analysis 
# tab before applying the Processed Data operator. It also changes the 
# analysis threshold from -80 dB to -65 dB in the 67 kHz processed data
# variable, in preparation for the school detection.

###################################################################################
# Required software:
# R
# Echoview

# Structure of folders should be:
# Instrument # \ Month \ files per day and frequency
# Folders should be named as follows:
# Instrument folder: e.g. 55026 (serial number)
# Month folder: e.g. 201505 for May 2015

# Shani Rousseau
# November 2016
#####################################################################################

###############################################
### -------------- EDIT ------------------- ###
###############################################
# ENTER SOURCE PATH
source_path <- "D:/ACRDP/2016/AZFP/55086/EV"

################################################
# --------------- DO NOT EDIT ---------------- #
################################################

# Define number of operational frequencies
numFreqs <- 4

# Define frequency vector
freqs <- c("67", "125", "200", "455")

library(RDCOMClient)

# Move into working directory
setwd(source_path)

# List all EV files
EVfiles.list <- list.files(source_path,pattern = ".*ed04_reviewed.EV$") 

L <- length(EVfiles.list)

for(i in 1:L){
  # Define EV filename
  EVfilename <- file.path(source_path,EVfiles.list[[i]])
  
  # Establish communication with Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # Open EV file
  EVfile <- EVApp$OpenFile(EVfilename)
              
  # Create editable line at 5m
  EVfile[["Lines"]]$CreateFixedDepth(5)
  
  # Rename line
  lineObj <- EVfile[["Lines"]]$FindByName('Line3')
  lineObj[["name"]] <- '5m'
  
  for(f in 1:numFreqs){
    # Exclude data above 5m and below bubbles line
    varName <- paste("Background noise removed Sv ",freqs[f]," kHz",sep="")
    Obj <- EVfile[["Variables"]]$FindByName(varName)$AsVariableAcoustic()
    Obj1 <- Obj[["Properties"]][["Analysis"]]
    Obj1[["ExcludeAboveLine"]] <- "5m"
    Obj1[["ExcludeBelowLine"]] <- "Bubbles"
  }
  
  # Apply -67 dB threshold to 67 kHz smoothed variable
  # Note that I am using -67 dB so that the school detection algorithm 
  # doesn't catch on the side lobes noise.
  Obj <- EVfile[["Variables"]]$FindByName("67 kHz smoothed")$AsVariableAcoustic()
  Obj1 <- Obj[["Properties"]][["Data"]]
  Obj1[["LockSvMinimum"]] <- FALSE
  Obj1[["ApplyMinimumThreshold"]] <- TRUE
  Obj1[["MinimumThreshold"]] = -67
  
  # Run school detection 
  schoolDet <- Obj$DetectSchools("Unclassified")
  if(schoolDet == -1){
    paste(EVfilename," : School detection failed.",sep="")
  }
  
  # Remove the -67 dB threshold
  Obj1 <- Obj[["Properties"]][["Data"]]
  Obj1[["ApplyMinimumThreshold"]] <- FALSE
  
  # Save new EV file as ed05_school_detection.EV
  nEVfilename <- gsub("ed04_reviewed","ed05_school_detection",EVfilename) 
  
  # Save file
  EVfile$SaveAs(nEVfilename)
  
  # Quit Echoview
  EVApp$Quit()
  
}  
# END OF FUNCTION
}
  