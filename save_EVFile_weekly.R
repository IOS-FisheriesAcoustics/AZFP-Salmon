save_EVFile_weekly <- function(){

require(RDCOMClient)
  
#############################################################################################
# This script generates weekly EV files containing AZFP .RAW data. A constant weekly pattern
# used (May 1-7, May 8-14 etc.) to allow for comparison between datasets. As a result, the 
# first and last EV files may not contain full weeks.

# Required software:
# R
# Echoview

# Structure of folders should be:
# Instrument # DATA \ Month \ raw files (monthly)
# Folders should be named as follows:
# Instrument folder: e.g. 55026 (serial number)
# Month folder: e.g. 201505 for May 2015
# E.g.: 55084\201505\rawFiles

# Shani Rousseau, M.Sc.
# shani.rousseau@gmail.com
# November 2016
  
  ###############################################
  ### -------------- EDIT ------------------- ###
  ###############################################
  
# ENTER SOURCE PATH. e.g. "E:/ACRDP/2015/AZFP/RAW/55086/DATA"
source.path <- "H:/ACRDP/2017/AZFP/55086/DATA"
# CREATE DESTINATION FOLDER
dest.path <- "D:/ACRDP/2017/AZFP/55086/EV"
# ENTER PATH To ECHOVIEW TEMPLATE
EVTemplate <- "D:/ACRDP/2017/EVTemplates/AZFPTemplate_2017.EV"
# ENTER PATH TO CALIBRATION FILES
cal.path <- "D:/ACRDP/2017/AZFP/55086/Calibration"
# NUMBER OF DAYS IN EACH EV FILE
dEV <- 7
# ENTER THE SOURCE FOLDER OF THE GPS TRACK
gpsFilename <- "D:/ACRDP/2017/AZFP/GPS_track/AZFP2017_GPS_track_1knot.gps.csv"

### ---------------------------------------- ###
################################################
### ---------------------------------------- ###


################################################
# --------------- DO NOT EDIT ---------------- #
################################################


# Extract instrument serial number
s <- regexpr("55",dest.path)
s.start <- s[[1]]
s.end <- s.start+4
inst <- substr(dest.path,s.start,s.end)

################################################
#          Create list of raw files            #
################################################

# Create string vector of all possible months
months.str <- c("May","June","July","August","September","October")

# List all month folders in the directory
monthList <- list.files(source.path, pattern = "^20")

# Define number of monthly folders
Lmonth = length(monthList)
months.str <- months.str[1:Lmonth]

# Create empty vector of length Lmonth
n <- vector("numeric", length = Lmonth)

# Create list of raw files (all season)
for(i in 1:Lmonth){
  Raw.files <- list.files(file.path(source.path,monthList[i]),pattern = "A$")
  Raw.path <- list.files(file.path(source.path,monthList[i]),pattern = "A$",full.names = TRUE)
  # Define number of hourly file for each month
  n[i] = length(Raw.path)
  
  if(i==1){
    # Define year, first month, first day of sampling
    year <- as.numeric(substr(Raw.files[[1]],1,2))
    fmonth.str <- substr(Raw.files[[1]],3,4)
    fmonth.num <- as.numeric(fmonth.str)
    fday <- as.numeric(substr(Raw.files[[1]],5,6))
    
    path.list <- Raw.path
  }else{
    path.list <- append(path.list,Raw.path,after = length(path.list))
  }
}

# Define total number of hourly files
nfiles = sum(n)


#######################################################
#     Create list of calibration files                #
#######################################################

calFiles.list <- list.files(cal.path,pattern="ecs$")

#######################################################
#      Create the first EV file (partial week)        #
#######################################################

# If first day of the first file is between 1 and 7, call it 1. If it falls between 8 and 14, call it 2. 
# If it falls between 15 and 21, call it 3. If it falls between 22 and 28, call it 4.
if(fday < 8){
  we.start <- 1
  #d.start <- 1
  #d.end <- 7
  d.vec <- c(1:7)
} else if(fday < 15){
  we.start <- 2
  #d.start <- 8
  #d.end <- 14
  d.vec <- c(8:14)
} else if(fday < 22){
  we.start <- 3
  #d.start <- 15
  #d.end <- 21
  d.vec <- c(15:21)
} else if(fday < 29){
  we.start <- 4
  #d.start <- 22
  #d.end <- 28
  d.vec <- c(22:28)
}

###########################################
#      Find files for first EV file       #
###########################################

# Generate reg expr that defines string to match
d.vec.str <- sprintf("%02d",d.vec)
d.vec.expr <- paste(d.vec.str,collapse="|")
expr <- paste(year,fmonth.str,"(",d.vec.expr,")",sep="")

# Match string
f <- grep(expr, path.list)
# List raw files
EV.list <- path.list[f]

################################################
#      Load these files into EV template       #
################################################

# Establish COM connection with Echoview
EVApp <- COMCreate("EchoviewCom.EvApplication")

# Open EV template in Echoview
EVFile <- EVApp$OpenFile(EVTemplate)

# Set fileset object
filesetObj <- EVFile[["Filesets"]]$Item(0)

# Load files
for(i in 1:length(EV.list)){
  filesetObj[["DataFiles"]]$Add(EV.list[[i]])
}

# Find month of middle raw file
LEV <- floor(length(EV.list)/2)
middle.file <- EV.list[LEV]
middle.month <- substring(sub(".*DATA/20[0-9]{2}","",middle.file),1,2)
#middle.month <- grep("DATA/20[0-9]{2}([0-9]{2})",middle.file,value=TRUE)

# Find calibration file
date <- paste("20",year,middle.month,sep="")
cal.file <- calFiles.list[grep(date,calFiles.list)]
cal.pathfile <- file.path(cal.path,cal.file)

# Add calibration file (month specific)
add.calibration <- filesetObj$SetCalibrationFile(cal.pathfile)

# Load GPS track - already in AZFPTemplate_2017.EV
#filesetObj <- EVFile[["Filesets"]]$Item(1)
#filesetObj[["DataFiles"]]$Add(gpsFilename)

# Save EV file
fullyear <- paste("20",year,sep="")
week.str <- paste("0",we.start,sep="")
week <- paste("Week",week.str,sep="")
filename <- paste(inst,fullyear,week,"ed00.EV",sep="_")
EVFile$SaveAS(file.path(dest.path ,filename))

# Close EV file
EVApp$CloseFile(EVFile)

# quit echoview
EVApp$Quit()

######################################################
#             Create remaining EV files              #
######################################################

# Define the number of remaining EV files to create
num.rfiles <- length(path.list)-length(EV.list)
nweek <- ceiling(num.rfiles/(7*24))


for(k in 1:nweek){
  
  # Define week number
  we.start <- we.start+1
  we.str <- sprintf("%02d",we.start)
  week <- paste("Week",we.str,sep="")
  
  # Last file of previous week
  last.file <- EV.list[length(EV.list)]

  # List files for this week
  loc.start <- grep(last.file,path.list)+1
  loc.end <- loc.start+(7*24)-1
  EV.list <- path.list[c(loc.start:loc.end)]
  
  ################################################
  #      Load these files into EV template       #
  ################################################
  
  # Establish COM connection with Echoview
  EVApp <- COMCreate("EchoviewCom.EvApplication")
  
  # Open EV template in Echoview
  EVFile <- EVApp$NewFile(EVTemplate)
  
  # Set fileset object
  filesetObj <- EVFile[["Filesets"]]$Item(0)
  
  # Load files
  for(i in 1:length(EV.list)){
    filesetObj[["DataFiles"]]$Add(EV.list[[i]])
  }
  
  # Find month of middle raw file
  LEV <- floor(length(EV.list)/2)
  middle.file <- EV.list[LEV]
  middle.month <- substring(sub(".*DATA/20[0-9]{2}","",middle.file),1,2)
  
  # Find calibration file
  date <- paste("20",year,middle.month,sep="")
  cal.file <- calFiles.list[grep(date,calFiles.list)]
  cal.pathfile <- file.path(cal.path,cal.file)
  
  # Add calibration file (month specific)
  add.calibration <- filesetObj$SetCalibrationFile(cal.pathfile)
  
  # Load GPS track
  filesetObj <- EVFile[["Filesets"]]$Item(1)
  filesetObj[["DataFiles"]]$Add(gpsFilename)
  
  # Save EV file
  filename <- paste(inst,fullyear,week,"ed00.EV",sep="_")
  EVFile$SaveAS(file.path(dest.path ,filename))
  
  # Close EV file
  EVApp$CloseFile(EVFile)
  
  # quit echoview
  EVApp$Quit()
  
}
# END OF FUNCTION
}
  
