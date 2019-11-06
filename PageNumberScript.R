# Set location of Data
setwd("E:/Voter")

# Import library for reading data and manipulation
library(openxlsx)
library(tidyverse)

# Import data
df1 <- read.xlsx("guj_ac_001.xlsx")

# Check length of the data
dim(df1)

# Sort data
df1 <- df1[order(df1$PartNo, df1$SectionNo, df1$SNo), ]

# Remove duplicates
df1 <- unique(df1)

# Check Number of rows
X = dim(df1)[1]
X
# Function to insert page number, Part Number and Section wise

df1$PageNo <- 0

k = 1
PageNo = 0

for (j in 2:X){
  
  if (df1$PartNo[j]==df1$PartNo[j-1]){
    
    if (df1$SectionNo[j]==df1$SectionNo[j-1]){
      k = k+1
      df1$PageNo[j] <- PageNo + ceiling(k/30)
    }
    else{
      PageNo = df1$PageNo[j-1]
      df1$PageNo[j] <- PageNo + 1
      k = 1
    }
  }
  else{
    PageNo = 0
    df1$PageNo[j] <- 1
    k = 1
  }
}
# Export data in Excel

write.xlsx(df1, file = "guj_ac_001_page.xlsx")

view(df1)
