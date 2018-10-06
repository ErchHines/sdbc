library(xlsx)
library(lubridate)
library(dplyr)
library(data.table)

clients <- read.xlsx("~//All SBDC clients 2018.xls", sheetName = "Report")
jobs <- read.xlsx("~//Jobs.xls", sheetName = "Report")
investment <- read.xlsx("~//Capital Investment.xls", sheetName = "Report")

#Get today's date and time so we can determine how old a firm is
today <- as.Date(Sys.Date(), format = 'POSIXct')

#Get the firm's age and put it in a new column
clients$Firm.Age <- 
  as.numeric(
    difftime(today, clients$Date.Company.Established, units = "weeks")
  )/52.25

#Change the NAs to 0
clients$Firm.Age[is.na(clients$Firm.Age)] <- 0

#This column provides us with the points for firm age, Note: The table is
#confusing since it goes from 1 to 3 then 4 to 7, what do you with firms that 
#are like 3 years and 8 months old? 

clients$Age.Points <- cut(
  clients$Firm.Age,
  breaks = c(-1, 1, 4, 8, Inf),
  labels = c(0, 5, 15, 25)
)

# Cnvert to numeric
clients$Age.Points <- as.numeric(as.character(clients$Age.Points))

#Same but for revenue
clients$Revenue.Points <- cut(
  clients$Gross.Revenue,
  breaks = c(-1, 40000, 125000, 1000000, Inf),
  labels = c(0, 5, 15, 25)
)

clients$Revenue.Points <- as.numeric(as.character(clients$Revenue.Points))

#Remove the NAs
clients$Revenue.Points[is.na(clients$Revenue.Points)] <- 0

#Create column for Industry.Points
clients$Industry.Points <- 0

#The following is assigning points based on industry, can't tell much from the cheat 
#get clarification on which categories get which points
clients[clients$Business.Type %in% 
  c('Accommodation/Food Svc.','Administrative/Support',
    'Educational', 'Health Care','Information',
    'Professional/Technical', 'Research and Development',
    'Service Establishment','Technology'),]$Industry.Points <- 5

clients[clients$Business.Type %in% 
  c('Arts and Entertainment','Real Estate',
    'Retail Dealer'),]$Industry.Points <- 15

clients[clients$Business.Type %in% 
  c('Agriculture','Construction Concern',
    'Manufacturer or Producer','Mining',
    'Transportation/Warehousing','Utilities',
    'Waste Management', 'Wholesale Dealer'),]$Industry.Points <- 25

#Create column for Referral.Points

clients$Referral.Points <- 0

#Again, get clarification on these categories

clients[clients$Referral.From %in% 
  c('Advertising/Marketing','Business Owner',
    'Internet','Lender','Newspapers/Magazines',
    'Other Client', 'Program Partners',
    'Training Seminar','Word of Mouth'),]$Referral.Points <- 25

#Add up the points
clients$Total.Points <- clients$Age.Points + clients$Revenue.Points +
  clients$Industry.Points + clients$Referral.Points


#Convert points to hours, again get clarification on this since I can't read
#the table in the picutre very well and they don't appear continuous
#What does 'Score Refer' Mean, gonna put 0 for now levels1 is startups and
#levels2 is everything else
levels1 <- c(-1, 20, 60, 80, Inf)
levels2 <- c(-1, 20, 40, 60, Inf)


labels <- c(1, 2, 3, 4)


clients$Hours <- with(clients, 
  ifelse(Initial.Business.Status %in% c("Pre-venture/Nascent", "Start-up (in bus. < 1 year)"),
         cut(clients$Total.Points, levels1, labels = labels),
         cut(clients$Total.Points, levels2, labels = labels)
         )
  )

#Note for me, figure out why it wouldn't work with levels and just returned
#number of the factors, in this case simple math will fix it though
clients$Hours <- (clients$Hours * 5) - 5

#Create a variable that gives how long the firm has been a client
clients$Client.Since <- 
  as.numeric(
    difftime(today, clients$Client.Signature.Start.Date, units = "weeks")
  )/52.25

#Assuming a new clients is a client that has been signed a year or less
clients$New.Client.Hours <- ifelse(clients$Client.Since <= 1, clients$Hours, 0)
clients$Existing.Hours <- ifelse(clients$Client.Since > 1, clients$Hours, 0)

#Change factirs to numbers
jobs$Milestone.Amount <- as.numeric(as.character(jobs$Milestone.Amount))
jobs$Change <- as.numeric(as.character(jobs$Change))

jobs[jobs$Type %in% "NewBus",]$Change <- 1

#Cast out miltestones and change numbers into seperate tables
milestones <- dcast(jobs, Client.ID ~ Type, sum, value.var =  "Milestone.Amount")
change <- dcast(jobs, Client.ID ~ Type, sum, value.var = "Change")

#Get the number of reports by client
reports <- as.data.frame(table(jobs$Client.ID))

#Merge in the number of reports into the clients database
clients <- merge(
  x = clients, y = reports,
  by.x = "Client.ID", by.y = "Var1", all.x = TRUE
)

#Change the name of the column to Reports and remove NAs
names(clients)[names(clients) == "Freq"] <- "Reports"
clients$Reports[is.na(clients$Reports)] <- 0

#Merge in investment data for amount approved
clients <- merge(
  x = clients, y = investment[c("Client.ID","Amount.Approved")],
  by = "Client.ID", all.x = TRUE
)

#If number of hours equals 0 report just the amount, look and see if this 
#goes away once we get a better idea of how hours are calculated
clients$Amount.Per.Hour <- ifelse(
  clients$Hours == 0, clients$Amount.Approved, 
  clients$Amount.Approved / clients$Hours
)

# Merge in the change table
clients <- merge(
  x = clients, y = change,
  by = "Client.ID", all.x = TRUE
)

#Get the total number of jobs added
clients$ChgStaff.All <- clients$`ChgStaff FT` + clients$`ChgStaff PT`

#Again, check to see what to do when hours = 0
clients$Sales.Per.Hour <- ifelse(
  clients$Hours == 0, clients$ChgSales,
  clients$ChgSales / clients$Hours
)

#Get business starts per hour
clients$NewBus.Per.Hour <- ifelse(
  clients$Hours == 0, clients$NewBus,
  clients$NewBus / clients$Hours
)

#Milestones per hour
clients$Milestones.Per.Hour <- ifelse(
  clients$Hours == 0, clients$Reports,
  clients$Reports / clients$Hours
)

#Get Business Start Contribution
clients$Buesiness.Start.Contrib <- clients$NewBus / sum(clients$NewBus, na.rm = TRUE)

clients$Buesiness.Start.Contrib <- clients$NewBus / sum(clients$NewBus, na.rm = TRUE)

