Consolidate<-function(directory=character()) {  #directory for folders, filename for individual files
  require(xlsx)
  files <- list.files(directory, pattern = "*.xlsx", full.name=T)
  #first create an empty data frame that we will insert data into later
  df_master <- data.frame(rptdate=as.Date(character()), rpttype=as.character(character()),  
                          va_ord_tot=as.numeric(character()), va_ord_pas=as.numeric(character()), 
                          dod_ord_tot=as.numeric(character()),dod_ord_pas=as.numeric(character()),
                          va_res_tot=as.numeric(character()), va_res_pas=as.numeric(character()), 
                          dod_res_tot=as.numeric(character()),dod_res_pas=as.numeric(character()))                          
  n=1
  for (row in files){
  #use For loop to look through every single file starting the first file
  filename = files [n] 
  rpttype <- c('')
  rowIndex <- numeric(0)
  if (grepl(" Con ", filename)) {
    rpttype<-c("Con")
    rowIndex <- c(7,8,11,12)
    rowDate <- (1)
  }
  else if (grepl(" Lab ", filename)) {
    rpttype<-c("Lab")
    rowIndex <- c(9,10,13,14,21,22,25,26)
    rowDate <- (1)    
  }
  else if (grepl(" Rad ", filename)) {
    rpttype<-c("Rad")
    rowIndex <- c(7,8,11,12)
    rowDate <- (2)
  }
  #get the files' initial day of the week ("day of the week" indicated in the first column) 
  rptdate <- read.xlsx(filename, sheetIndex=2, colIndex=3, rowIndex=rowDate, header=F, colClasses="Date")
  day <- weekdays(as.Date(rptdate[1,1])) #convert date to day of the week
  
  #if the initial day is Sunday then we set i = 2 so we can read the next 2 columns which are Saturday and Friday
  if (day == "Sunday") {
    i = 2
  }
  else{
    i = 0
  }
  for (k in 0:i){
    rptdate <- read.xlsx(filename, sheetIndex=2, colIndex=(3+k), rowIndex=rowDate, header=F, colClasses='Date')
    totals <- read.xlsx(filename, sheetIndex=2, colIndex=(3+k), rowIndex=rowIndex, header=F, colClasses="numeric")
    df <- data.frame(rptdate=as.Date(rptdate[1,1]), rpttype=as.character(rpttype),  
          va_ord_tot=as.numeric(totals[1,1]), va_ord_pas=as.numeric(totals[2,1]), 
          dod_ord_tot=as.numeric(totals[3,1]),dod_ord_pas=as.numeric(totals[4,1]),
          va_res_tot=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[5,1]), 
          va_res_pas=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[6,1]), 
          dod_res_tot=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[7,1]), 
          dod_res_pas=ifelse(rpttype=='Con'|rpttype=='Rad', NA, totals[8,1]))
    df$rpttype <- as.character(df$rpttype)  
    df_master <- rbind(df_master, df) #use rbind to combine all data frames  
    }
    n=n+1     
  }
  df_master <- df_master[order(as.Date(df_master$rptdate)), ] #order by rptdate
  print(df_master, quote = FALSE, row.names = FALSE)
  #summary(df_master$va_ord_tot+df_master$dod_ord_tot)
  #hist(df_master$va_ord_tot+df_master$dod_ord_tot)
  #write.csv(df_master, file="Combined.csv")
  #sum(df_master$va_ord_tot, df_master$va_ord_pas)
}
