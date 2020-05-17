# preparing the staging file for Rollup

getwd()
setwd("C:/Users/WPKH43/Desktop/D-Top Folder/Rollup Automation Project")

install.packages("xlsx")
library(xlsx)                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                                          

install.packages("tm")
library(tm)

install.packages("tidyr")
library(tidyr)

install.packages("dplyr")
library(dplyr)

install.packages("Hmisc")
library(Hmisc)

install.packages("stringr")
library(stringr)

install.packages("sos")
library(sos)

#Fetching the SIP Matrix for SIP Lookup
#The file should have the following headers in Sheet1 Tab
#---------------------------------------------------------------------------------------------------------
#Employee Id,	Person Name,	First Name,	Last Name,	Manager, 	SFDC Region,	SFDC Subregion,	SFDC Position
#---------------------------------------------------------------------------------------------------------
SIP.LookupSheet <- read.xlsx('Master Copy of SIP.xlsx', sheetName = 'Sheet1', header = TRUE, stringsAsFactors = FALSE)

#Cleanse and Combine First.Name and Last.Name column into 1 column Full.Name for SIP Lookupsheet
#Removing next line character 
SIP.LookupSheet$First.Name<-gsub("[\n]", "", SIP.LookupSheet$First.Name)
SIP.LookupSheet$Last.Name<-gsub("[\n]", "", SIP.LookupSheet$Last.Name)

#Removing leading and trailing spaces
SIP.LookupSheet$First.Name<-trimws(SIP.LookupSheet$First.Name)
SIP.LookupSheet$Last.Name<-trimws(SIP.LookupSheet$Last.Name)

#Make all charaters to lower case
SIP.LookupSheet$First.Name<-tolower(SIP.LookupSheet$First.Name)
SIP.LookupSheet$Last.Name<-tolower(SIP.LookupSheet$Last.Name)
SIP.LookupSheet$Employee.Id<-tolower(SIP.LookupSheet$Employee.Id)

#Combining First.Name and Last.Name into Full.Name
SIP.LookupSheet$Full.Name<-NULL
SIP.LookupSheet$Full.Name<-paste(SIP.LookupSheet$First.Name,SIP.LookupSheet$Last.Name)

#Cleansing 'Employee.Id' column in SIP Lookup
colnames(SIP.LookupSheet)[colnames(SIP.LookupSheet)=="Employee.Id"] <- "PersonID"
#Removing next line character 
SIP.LookupSheet$PersonID<-gsub("[\n]", "", SIP.LookupSheet$PersonID)

#Removing leading and trailing spaces
SIP.LookupSheet$PersonID<-trimws(SIP.LookupSheet$PersonID)

#Fetching the ProductLine Code for product Line Lookup
#The file should have the following headers in Sheet1 Tab
#---------------------------------------------------------------------------------------------------------
#Code	ProductLine
#---------------------------------------------------------------------------------------------------------
ProdLTbl <- read.csv('ProductLineCode.csv', header = TRUE, stringsAsFactors = FALSE)

ProdLTbl$Code<-trimws(ProdLTbl$Code)
ProdLTbl$ProductLine<-trimws(ProdLTbl$ProductLine)

ProdLTbl$Code<-tolower(ProdLTbl$Code)
ProdLTbl$ProductLine<-tolower(ProdLTbl$ProductLine)

#convert all the codes in a vector to compare
ProdLineVector <- ProdLTbl$Code

#Fetching the ProductSubset Code for product subset Lookup
#The file should have the following headers in Sheet1 Tab
#---------------------------------------------------------------------------------------------------------
#Code	ProductSubset
#---------------------------------------------------------------------------------------------------------
ProdSTbl <- read.csv('ProductSubsetCode.csv', header = TRUE, stringsAsFactors = FALSE)

ProdSTbl$Code<-trimws(ProdSTbl$Code)
ProdSTbl$ProductSubset<-trimws(ProdSTbl$ProductSubset)

ProdSTbl$Code<-tolower(ProdSTbl$Code)
ProdSTbl$ProductSubset<-tolower(ProdSTbl$ProductSubset)

#convert all the codes in a vector to compare
ProdSubVector <- ProdSTbl$Code

#Fetching the ProductSubset Code for product subset Lookup
#The file should have the following headers in Sheet1 Tab
#---------------------------------------------------------------------------------------------------------
#Code	SolutionFocusCode
#---------------------------------------------------------------------------------------------------------
SolFTbl <- read.csv('SolutionFocusCode.csv', header = TRUE, stringsAsFactors = FALSE)

SolFTbl$Code<-trimws(SolFTbl$Code)
SolFTbl$SolnFocus<-trimws(SolFTbl$SolnFocus)

SolFTbl$Code<-tolower(SolFTbl$Code)
SolFTbl$SolnFocus<-tolower(SolFTbl$SolnFocus)

#convert all the codes in a vector to compare
SolFVector <- SolFTbl$Code

#Fetching the Sourcesystem Code for Source Lookup
#The file should have the following headers in Sheet1 Tab
#---------------------------------------------------------------------------------------------------------
#Code	Source
#---------------------------------------------------------------------------------------------------------
SourceSTbl <- read.csv('SourceSystemCode.csv', header = TRUE, stringsAsFactors = FALSE)

SourceSTbl$Code<-trimws(SourceSTbl$Code)
SourceSTbl$Source<-trimws(SourceSTbl$Source)

SourceSTbl$Code<-tolower(SourceSTbl$Code)
SourceSTbl$Source<-tolower(SourceSTbl$Source)

#convert all the codes in a vector to compare
SourceSystemVector <- SourceSTbl$Code

#START: Code added on MAY 2020#
# Note: This code is added to further enhance the Rollup script.
#       Raw SIP file can pe given as input to this one
#Fetching the raw SIP Matrix 
#The file should have the following headers
#--------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
#Employee ID	Person Name	First Name	Last Name	Manager	CK Direct Reports	SFDC Region	SFDC Subregion	SFDC Position	LEVEL	SEQUENCE	Plan Type 2020	Metric 1	Metric 2	Metric 3
#---------------------------------------------------------------------------------------------------------------------------------------------------------------------------------
SIP.Matrix <- read.xlsx('InputRawSIP.xlsx', sheetName = 'Sheet1', header = TRUE, stringsAsFactors = FALSE)

#fetching only useful columns
# Note: select() function is part of dplyr package
SIP.Matrix<-select(SIP.Matrix,"Employee.ID","Person.Name","SFDC.Position","SEQUENCE","Metric.1","Metric.2","Metric.3")



# Find the test Rollup in Metric 1
SIP.Matrix1<-filter(SIP.Matrix, grepl(c("rollup|rollup:"),SIP.Matrix$Metric.1, ignore.case = TRUE))
SIP.Matrix1$Metric.2<-NULL
SIP.Matrix1$Metric.3<-NULL
SIP.Matrix1$MetricTo<-"Rev 1"
colnames(SIP.Matrix1)[colnames(SIP.Matrix1)=="Metric.1"]<-"Rollup"


# Find the test Rollup in Metric 2
SIP.Matrix2<-filter(SIP.Matrix, grepl(c("rollup|rollup:"),SIP.Matrix$Metric.2, ignore.case = TRUE))
SIP.Matrix2$Metric.1<-NULL
SIP.Matrix2$Metric.3<-NULL
SIP.Matrix2$MetricTo<-"Rev 2"
colnames(SIP.Matrix2)[colnames(SIP.Matrix2)=="Metric.2"]<-"Rollup"

# Find the test Rollup in Metric 3
SIP.Matrix3<-filter(SIP.Matrix, grepl(c("rollup|rollup:"),SIP.Matrix$Metric.3, ignore.case = TRUE))
SIP.Matrix3$Metric.1<-NULL
SIP.Matrix3$Metric.2<-NULL
SIP.Matrix3$MetricTo<-"Rev 3"
colnames(SIP.Matrix3)[colnames(SIP.Matrix3)=="Metric.3"]<-"Rollup"

# combining all
SIP.trans<-rbind(SIP.Matrix1, SIP.Matrix2, SIP.Matrix3)

# Updating sequence
# finding sequence with , in them and without , in them
# Where without , is found duplicate the same seq and add ,
# this is use to create a standard format of sequence and then split it into 3 metrics
SIP.trans$Seqdiff<-ifelse((grepl(c(","),SIP.trans$SEQUENCE, ignore.case = TRUE) ),"TRUE","FALSE")

SIP.trans$SEQUENCE1<-ifelse((SIP.trans$Seqdiff==FALSE), paste(SIP.trans$SEQUENCE,",",SIP.trans$SEQUENCE,",",SIP.trans$SEQUENCE, sep=""), SIP.trans$SEQUENCE)

#REmove the blank spaces
SIP.trans$SEQUENCE1<- trimws(SIP.trans$SEQUENCE1)

SIP.trans<-separate(data=SIP.trans, col=SEQUENCE1, into=c("M1seq","M2seq","M3seq"))

SIP.trans$SEQUENCE2<-ifelse((SIP.trans$MetricTo=='Rev 1'), SIP.trans$M1seq, 
                            ifelse((SIP.trans$MetricTo=='Rev 2'), SIP.trans$M2seq, 
                                   ifelse((SIP.trans$MetricTo=='Rev 3'), SIP.trans$M3seq,"")))
#write.csv(SIP.trans, file = "SIPRaw.csv", row.names = FALSE, na="")

#Renaming the file and columns to Match input for already written code
SIP.RawFile<-SIP.trans

colnames(SIP.RawFile)[colnames(SIP.RawFile)=="Employee.ID"] <- "MngID"
colnames(SIP.RawFile)[colnames(SIP.RawFile)=="Person.Name"] <- "MngName"
colnames(SIP.RawFile)[colnames(SIP.RawFile)=="SFDC.Position"] <- "MngPosition"
colnames(SIP.RawFile)[colnames(SIP.RawFile)=="SEQUENCE2"] <- "Sq"

#dropping unneccessary columns
SIP.RawFile$SEQUENCE<-NULL
SIP.RawFile$Seqdiff<-NULL
SIP.RawFile$M1seq<-NULL
SIP.RawFile$M2seq<-NULL
SIP.RawFile$M3seq<-NULL

#adjust Column positions
SIP.RawFile<- SIP.RawFile[c(1,2,3,5,4,6)]

#END: Code added on MAY 2020#

SIP.RawFile$Rollup<-gsub("=","",SIP.RawFile$Rollup)

#Removing Trailing and leading zeros
SIP.RawFile$Rollup<-trimws(SIP.RawFile$Rollup)
# make all lower case
SIP.RawFile$Rollup<-tolower(SIP.RawFile$Rollup)

# Breaking Down SIP Metric Sheet into Rollup
##SIP.RawFile<- separate_rows(SIP.RawFile, Rollup,sep = '&')
SIP.RawFile<- separate_rows(SIP.RawFile, Rollup,sep = '\\+')

#Finding if the Rollup has multiple metrics  
#Finding by searching for '&' sign. following is the function
#SIP.RawFile$MultipleMtrType<-ifelse(grepl("&",SIP.RawFile$Rollup),"TRUE","FALSE")


#Check if the metric has metric 1
SIP.RawFile$Metric1Found<-ifelse((grepl(c("metric 1|metric1|metrics 1|metrics1"),SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")
#Check if the metric has metric 2
SIP.RawFile$Metric2Found<-ifelse((grepl(c("metric 2|metric2|metrics 2|metrics2"),SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")
#Check if the metric has metric 3
SIP.RawFile$Metric3Found<-ifelse((grepl(c("metric 3|metric3|metrics 3|metrics3"),SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")

#Check if the metric has 'AIT HW' in it
##########JAN 2019
SIP.RawFile$AITFound<-ifelse((grepl(c("AIT HW|ait hw|AIT Service Bookings|ait service bookings"),SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")


#Check if the metric has 'Bookings' in it
SIP.RawFile$ServicesFound<-ifelse((grepl(c("booking|bookings|service|services"),SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")

# fetching all employee ids
# fetching Employee ID from 'Person.Name' column
SIP.RawFile$PersonID<-NULL
#SIP.RawFile$PersonID<-ifelse( grepl("(",SIP.RawFile$Rollup,fixed=TRUE), gsub(".*\\(\\s*|\\).*", "",SIP.RawFile$Rollup),"")

SIP.RawFile$PersonID<-gsub("[\\(\\)]", "", regmatches(SIP.RawFile$Rollup, gregexpr("\\(.*?\\)",SIP.RawFile$Rollup)))

#Cleaning up the PersonID
SIP.RawFile$PersonID<-gsub("[\n]", "", SIP.RawFile$PersonID)

SIP.RawFile$PersonID<-gsub("\"", "", SIP.RawFile$PersonID)

SIP.RawFile$PersonID<-gsub("c", "", SIP.RawFile$PersonID)

SIP.RawFile$PersonID<-gsub(",", ";", SIP.RawFile$PersonID)

SIP.RawFile$Rollup.Staging<-NULL
SIP.RawFile$Rollup.Staging1<-NULL

#####################2020February#############################
##Start Changes

SIP.RawFile$Rollup<-gsub("solall", "solf1 solf2 solf3 solf4 solf5 solf6 solf7 solf8 solf9 solf10 solf11", SIP.RawFile$Rollup)

##End Changes
#######################################

###################2019DEC#########################
#Product Line cleanup
#Check if the metric has Product Line
SIP.RawFile$PLFound<-ifelse((grepl(ProdLineVector,SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")

match1<-NULL
match1<- lapply(SIP.RawFile$Rollup, function(x)
  str_extract(x,ProdLineVector))
SIP.RawFile$prodLineStaging<-NULL  
SIP.RawFile$prodLineStaging<- as.character(t(t(match1))) 

#extracting only prodline from the staging column
SIP.RawFile$prodLineStaging1 <- as.character(str_extract_all(SIP.RawFile$prodLineStaging, '".*"'))

#######2020 Feb changes for Prod Line
SIP.RawFile$prodLineStaging1<- as.character(str_replace_all(SIP.RawFile$prodLineStaging1, "NA,", ""))
#Removing spaces and cleanup
SIP.RawFile$prodLineStaging1<-gsub("[\r\n]", "", SIP.RawFile$prodLineStaging1)
SIP.RawFile$prodLineStaging1<-gsub("\"", "", SIP.RawFile$prodLineStaging1)
SIP.RawFile$prodLineStaging1<-gsub(" ", "", SIP.RawFile$prodLineStaging1)
SIP.RawFile$prodLineStaging1<-trimws(SIP.RawFile$prodLineStaging1)

#replacing all the character(0) with '0'
SIP.RawFile$prodLineStaging1[(SIP.RawFile$prodLineStaging1=='character(0)')]<-0

# removing "" and replacing | instead of ,
SIP.RawFile$prodLineStaging1<- gsub('"',"",SIP.RawFile$prodLineStaging1)
SIP.RawFile$prodLineStaging1<- gsub(",","|",SIP.RawFile$prodLineStaging1)



#Product SUbset cleanup
#Check if the metric has Product Subset
SIP.RawFile$PSSFound<-ifelse((grepl(ProdSubVector,SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")

match2<-NULL
match2<- lapply(SIP.RawFile$Rollup, function(x)
  str_extract(x,ProdSubVector))
SIP.RawFile$prodSubStaging<-NULL  
SIP.RawFile$prodSubStaging<- as.character(t(t(match2))) 

#extracting only prodline from the staging column
SIP.RawFile$prodSubStaging1 <- as.character(str_extract_all(SIP.RawFile$prodSubStaging, '".*"'))

#######2020 Feb changes for Prod Subset
SIP.RawFile$prodSubStaging1<- as.character(str_replace_all(SIP.RawFile$prodSubStaging1, "NA,", ""))
#Removing spaces and cleanup
SIP.RawFile$prodSubStaging1<-gsub("[\r\n]", "", SIP.RawFile$prodSubStaging1)
SIP.RawFile$prodSubStaging1<-gsub("\"", "", SIP.RawFile$prodSubStaging1)
SIP.RawFile$prodSubStaging1<-gsub(" ", "", SIP.RawFile$prodSubStaging1)
SIP.RawFile$prodSubStaging1<-trimws(SIP.RawFile$prodSubStaging1)

#replacing all the character(0) with '0'
SIP.RawFile$prodSubStaging1[(SIP.RawFile$prodSubStaging1=='character(0)')]<-0

# removing "" and replacing | instead of ,
SIP.RawFile$prodSubStaging1<- gsub('"',"",SIP.RawFile$prodSubStaging1)
SIP.RawFile$prodSubStaging1<- gsub(",","|",SIP.RawFile$prodSubStaging1)



#SOlution Focus cleanup
#Check if the metric has Solution Focus
SIP.RawFile$SOFFound<-ifelse((grepl(SolFVector,SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")

match2<-NULL
match2<- lapply(SIP.RawFile$Rollup, function(x)
  str_extract(x,SolFVector))
SIP.RawFile$SolFocusStaging<-NULL  
SIP.RawFile$SolFocusStaging<- as.character(t(t(match2))) 

#extracting only Solution Focus from the staging column
SIP.RawFile$SolFocusStaging1 <- as.character(str_extract_all(SIP.RawFile$SolFocusStaging, '".*"'))

#######2020 Feb changes for SolFocus
SIP.RawFile$SolFocusStaging1<- as.character(str_replace_all(SIP.RawFile$SolFocusStaging1, "NA,", ""))
SIP.RawFile$SolFocusStaging1<- as.character(str_replace_all(SIP.RawFile$SolFocusStaging1, "NA", ""))
#Removing spaces and cleanup
SIP.RawFile$SolFocusStaging1<-gsub("[\r\n]", "", SIP.RawFile$SolFocusStaging1)
SIP.RawFile$SolFocusStaging1<-gsub("\"", "", SIP.RawFile$SolFocusStaging1)
SIP.RawFile$SolFocusStaging1<-gsub(" ", "", SIP.RawFile$SolFocusStaging1)
SIP.RawFile$SolFocusStaging1<-trimws(SIP.RawFile$SolFocusStaging1)

#replacing all the character(0) with '0'
SIP.RawFile$SolFocusStaging1[(SIP.RawFile$SolFocusStaging1=='character(0)')]<-0

# removing "" and replacing | instead of ,
SIP.RawFile$SolFocusStaging1<- gsub('"',"",SIP.RawFile$SolFocusStaging1)
SIP.RawFile$SolFocusStaging1<- gsub(",","|",SIP.RawFile$SolFocusStaging1)


###################################################
############APRIL 2020

#Source System cleanup
#Check if the metric has Source System
SIP.RawFile$SOSFound<-ifelse((grepl(SourceSystemVector,SIP.RawFile$Rollup, ignore.case = TRUE) ),"TRUE","FALSE")

match2<-NULL
match2<- lapply(SIP.RawFile$Rollup, function(x)
  str_extract(x,SourceSystemVector))
SIP.RawFile$SourceStaging<-NULL  
SIP.RawFile$SourceStaging<- as.character(t(t(match2))) 

#extracting only SOurce System from the staging column
SIP.RawFile$SourceStaging1 <- as.character(str_extract_all(SIP.RawFile$SourceStaging, '".*"'))

#######2020 APR changes for Source System
SIP.RawFile$SourceStaging1<- as.character(str_replace_all(SIP.RawFile$SourceStaging1, "NA,", ""))
SIP.RawFile$SourceStaging1<- as.character(str_replace_all(SIP.RawFile$SourceStaging1, "NA", ""))
#Removing spaces and cleanup
SIP.RawFile$SourceStaging1<-gsub("[\r\n]", "", SIP.RawFile$SourceStaging1)
SIP.RawFile$SourceStaging1<-gsub("\"", "", SIP.RawFile$SourceStaging1)
#SIP.RawFile$SourceStaging1<-gsub(" ", "", SIP.RawFile$SourceStaging1)
SIP.RawFile$SourceStaging1<-trimws(SIP.RawFile$SourceStaging1)

#replacing all the character(0) with '0'
SIP.RawFile$SourceStaging1[(SIP.RawFile$SourceStaging1=='character(0)')]<-0

# removing "" and replacing | instead of ,
SIP.RawFile$SourceStaging1<- gsub('"',"",SIP.RawFile$SourceStaging1)
SIP.RawFile$SourceStaging1<- gsub(",","|",SIP.RawFile$SourceStaging1)


###################################################

SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& (SIP.RawFile$AITFound==TRUE )& SIP.RawFile$ServicesFound==FALSE)]<-paste("Metric 1,AIT,,,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==FALSE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)]<-paste("Metric 1,,,,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)]<-paste("Metric 1,,,Services,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& (SIP.RawFile$AITFound==TRUE ) & SIP.RawFile$ServicesFound==TRUE)]<-paste("Metric 1,AIT,,Services,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==TRUE)],sep="")

SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==FALSE)]<-paste("Metric 2,AIT,,,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==FALSE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)]<-paste("Metric 2,,,,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)]<-paste("Metric 2,,,Services,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==TRUE)]<-paste("Metric 2,AIT,,Services,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==TRUE)],sep="")

SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==FALSE)]<-paste("Metric 3,AIT,,,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==FALSE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)]<-paste("Metric 3,,,,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)]<-paste("Metric 3,,,Services,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)],sep="")
SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==TRUE)]<-paste("Metric 3,AIT,,Services,,",SIP.RawFile$PersonID[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==TRUE)],sep="")

SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric1Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==FALSE)]<-gsub(";",";Metric 1,AIT,,,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==FALSE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)]<-gsub(";",";Metric 1,,,,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)]<-gsub(";",";Metric 1,,,Services,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric1Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==TRUE)]<-gsub(";",";Metric 1,AIT,,Services,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric1Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==TRUE)], ignore.case = TRUE)

SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric2Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==FALSE)]<-gsub(";",";Metric 2,AIT,,,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==FALSE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)]<-gsub(";",";Metric 2,,,,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)]<-gsub(";",";Metric 2,,,Services,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric2Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==TRUE)]<-gsub(";",";Metric 2,AIT,,Services,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric2Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==TRUE)], ignore.case = TRUE)

SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric3Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==FALSE)]<-gsub(";",";Metric 3,AIT,,,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==FALSE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)]<-gsub(";",";Metric 3,,,,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==FALSE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)]<-gsub(";",";Metric 3,,,Services,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==FALSE & SIP.RawFile$ServicesFound==TRUE)], ignore.case = TRUE)
SIP.RawFile$Rollup.Staging1[(SIP.RawFile$Metric3Found==TRUE& (SIP.RawFile$AITFound==TRUE) & SIP.RawFile$ServicesFound==TRUE)]<-gsub(";",";Metric 3,AIT,,Services,,",SIP.RawFile$Rollup.Staging[(SIP.RawFile$Metric3Found==TRUE& SIP.RawFile$AITFound==TRUE & SIP.RawFile$ServicesFound==TRUE)], ignore.case = TRUE)


#SIP.RawFile[c("MngID","MetricTo","MngName","MngPosition","Rollup.Staging1","Sq")]

SIP.Rollup.Stg <- SIP.RawFile[c("MngID","MetricTo","MngName","MngPosition","Rollup.Staging1","Sq","prodLineStaging1", "prodSubStaging1","SolFocusStaging1","SourceStaging1")]

# Breaking Down Metric 1 Staging Sheet into Rollup
Rollup.Final<-NULL
Rollup.Final<- separate_rows(SIP.Rollup.Stg, Rollup.Staging1,sep = ';')

#Changes to Include services LS MS PS 2019
#Rollup.Final<-separate_rows(Rollup.Final, ProdFam, sep="\\|")

Rollup.Final<-separate(Rollup.Final, Rollup.Staging1, into=c("MetricFrom","Source","Product SubSet","Product Line","Soln Focus","PersonID"), sep=",", remove=TRUE,convert=FALSE)

#################Jan 2020##########################
#Update productLine with staging where ever prod line is populated
Rollup.Final$`Product Line`<- ifelse(Rollup.Final$prodLineStaging1==0, Rollup.Final$`Product Line` ,Rollup.Final$prodLineStaging1)

#clean the productLine column for any space
Rollup.Final$`Product Line`<-trimws(Rollup.Final$`Product Line`)

#Removing next line character 
Rollup.Final$`Product Line`<-gsub("[\n]", "", Rollup.Final$`Product Line`)

#Separate Product Line column by |
Rollup.Final<-separate_rows(Rollup.Final, `Product Line`, sep="\\|")

#clean the productLine column for any space, after the split
Rollup.Final$`Product Line`<-trimws(Rollup.Final$`Product Line`)

#Removing next line character, after the split
Rollup.Final$`Product Line`<-gsub("[\n]", "", Rollup.Final$`Product Line`)

colnames(Rollup.Final)[colnames(Rollup.Final)=="Product Line"] <- "Code"

#Lookup the Product Line in the ProdLTbl
Rollup.Final<-  merge(Rollup.Final, ProdLTbl[,c("Code","ProductLine")], by="Code", all.x = TRUE)

Rollup.Final$Code<-NULL
Rollup.Final$prodLineStaging1<-NULL

#####################################
#Update productSubSet with staging where ever prod subset is populated
Rollup.Final$`Product SubSet`<- ifelse(Rollup.Final$prodSubStaging1==0, Rollup.Final$`Product SubSet` ,Rollup.Final$prodSubStaging1)

#clean the productSubset column for any space
Rollup.Final$`Product SubSet`<-trimws(Rollup.Final$`Product SubSet`)

#Removing next line character 
Rollup.Final$`Product SubSet`<-gsub("[\n]", "", Rollup.Final$`Product SubSet`)

#Separate Product Subset column by |
#Rollup.Final<-separate_rows(Rollup.Final, `Product Line`, sep="\\|")

#clean the productSubset column for any space, after the split
Rollup.Final$`Product SubSet`<-trimws(Rollup.Final$`Product SubSet`)

#Removing next line character, after the split
Rollup.Final$`Product SubSet`<-gsub("[\n]", "", Rollup.Final$`Product SubSet`)

colnames(Rollup.Final)[colnames(Rollup.Final)=="Product SubSet"] <- "Code"

#Lookup the Product Subset in the ProdLTbl
Rollup.Final<-  merge(Rollup.Final, ProdSTbl[,c("Code","ProductSubset")], by="Code", all.x = TRUE)

Rollup.Final$Code<-NULL
Rollup.Final$prodSubStaging1<-NULL

#####################################
#Update Solution Focus with staging where ever Solution Focus is populated
Rollup.Final$`Soln Focus`<- ifelse(Rollup.Final$SolFocusStaging1==0, Rollup.Final$`Soln Focus` ,Rollup.Final$SolFocusStaging1)

#clean the Solution Focus column for any space
Rollup.Final$`Soln Focus`<-trimws(Rollup.Final$`Soln Focus`)

#Removing next line character 
Rollup.Final$`Soln Focus`<-gsub("[\n]", "", Rollup.Final$`Soln Focus`)

#Separate Solution Focus column by |
Rollup.Final<-separate_rows(Rollup.Final, `Soln Focus`, sep="\\|")

#clean the Solution Focus column for any space, after the split
Rollup.Final$`Soln Focus`<-trimws(Rollup.Final$`Soln Focus`)

#Removing next line character, after the split
Rollup.Final$`Soln Focus`<-gsub("[\n]", "", Rollup.Final$`Soln Focus`)

colnames(Rollup.Final)[colnames(Rollup.Final)=="Soln Focus"] <- "Code"

#Lookup the Solution Focus in the SolFTbl
Rollup.Final<-  merge(Rollup.Final, SolFTbl[,c("Code","SolnFocus")], by="Code", all.x = TRUE)

Rollup.Final$Code<-NULL
Rollup.Final$SolFocusStaging1<-NULL

#####################################
#Update Source with staging where ever Source is populated
Rollup.Final$`Source`<- ifelse(Rollup.Final$SourceStaging1==0, Rollup.Final$`Source` ,Rollup.Final$SourceStaging1)

#clean the Source column for any space
Rollup.Final$`Source`<-trimws(Rollup.Final$`Source`)

#Removing next line character 
Rollup.Final$`Source`<-gsub("[\n]", "", Rollup.Final$`Source`)

#Separate Source column by |
Rollup.Final<-separate_rows(Rollup.Final, `Source`, sep="\\|")

#clean the Solution Focus column for any space, after the split
Rollup.Final$`Source`<-trimws(Rollup.Final$`Source`)

#Removing next line character, after the split
Rollup.Final$`Source`<-gsub("[\n]", "", Rollup.Final$`Source`)

colnames(Rollup.Final)[colnames(Rollup.Final)=="Source"] <- "Code"

#Lookup the Solution Focus in the SourceSFTbl
Rollup.Final<-  merge(Rollup.Final, SourceSTbl[,c("Code","Source")], by="Code", all.x = TRUE)

Rollup.Final$Code<-NULL
Rollup.Final$SourceStaging1<-NULL

###############################################################################

#Clean PersonID column
#Removing next line character 
Rollup.Final$ PersonID<-gsub("[\n]", "", Rollup.Final$ PersonID)

#Removing leading and trailing spaces
Rollup.Final$ PersonID<-trimws(Rollup.Final$ PersonID)


#Change Metric 1/2/3 to Rev 1/2/3
Rollup.Final$MetricFrom<-tolower(Rollup.Final$MetricFrom)
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metric 1"]<-"Rev 1"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metric1"]<-"Rev 1"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metrics 1"]<-"Rev 1"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metrics1"]<-"Rev 1"

Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metric 2"]<-"Rev 2"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metric2"]<-"Rev 2"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metrics 2"]<-"Rev 2"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metrics2"]<-"Rev 2"

Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metric 3"]<-"Rev 3"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metric3"]<-"Rev 3"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metrics 3"]<-"Rev 3"
Rollup.Final$MetricFrom[Rollup.Final$MetricFrom=="metrics3"]<-"Rev 3"


#Match based on Full Name for all the NA
Rollup.Final<-  merge(Rollup.Final, SIP.LookupSheet[,c("SFDC.Position","Person.Name","PersonID")], by="PersonID", all.x = TRUE)

#Renaming the Column to match the standard Rollup sheet
colnames(Rollup.Final)[colnames(Rollup.Final)=="PersonID"] <- "RepID"
colnames(Rollup.Final)[colnames(Rollup.Final)=="SFDC.Position"] <- "RepIDPosition"
colnames(Rollup.Final)[colnames(Rollup.Final)=="Person.Name"] <- "RepIDName"
colnames(Rollup.Final)[colnames(Rollup.Final)=="ProductLine"] <- "Product Line"
colnames(Rollup.Final)[colnames(Rollup.Final)=="ProductSubset"] <- "Product Subset"
colnames(Rollup.Final)[colnames(Rollup.Final)=="SolnFocus"] <- "Soln Focus"

#Converting the column RepID to Upper case for all the NATEMPs
Rollup.Final$RepID<-toupper(Rollup.Final$RepID)

#Rearranging the column positions to match the standard Rollup sheet
Rollup.Final<- Rollup.Final[c(2,1,6,3,11,9,8,10,4,5,13,12,7)]

#Renaming the Column 

write.csv(Rollup.Final, file = "MyData.csv", row.names = FALSE, na="")