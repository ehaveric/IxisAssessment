library(readr)

dailysesh <- read.csv("C:/Users/ehaveric99/Documents/technical exercises/DataAnalyst_Ecom_data_sessionCounts.csv")
monthlyadds <- read.csv("C:/Users/ehaveric99/Documents/technical exercises/DataAnalyst_Ecom_data_addsToCart.csv")

monthlyadds

install.packages('lubridate')
library(lubridate)
install.packages('tidyverse')
library(tidyverse)
library(zoo)

str(dailysesh$dim_date)
dailysesh$dim_date <- mdy(dailysesh$dim_date)

monthlysesh <- dailysesh
monthlysesh$year_month <- floor_date(monthlysesh$dim_date,unit = "month")

#The floor_date helps us create a new row that defaults all days
#in a month to the first day of the month, this will help when 
#aggregating the monthly data below.

library(dplyr)

monthlyseshagg <- monthlysesh %>%
  group_by(dim_browser, dim_deviceCategory, year_month) %>%
  dplyr::summarize(sessions=sum(sessions), 
                   transactions=sum(transactions), QTY=sum(QTY)) %>%
  as.data.frame()

##We now have monthly aggregated data for each classifier in the
##browser and device category columns.

#To add the ECR we simply divide transactions with sessions

monthlyseshagg$ECR <- monthlyseshagg$transactions /
  monthlyseshagg$sessions

#This creates some NA values when we have an indeterminate solution
#Can format for those as such.

monthlyseshagg["ECR"][is.na(monthlyseshagg["ECR"])] <- 0

#Since we need a month * device aggregation, we can keep this
#master aggregation in case we need it in the future, and remake
# a more concise one 

deviceagg <- monthlysesh %>%
  group_by(dim_deviceCategory, year_month) %>%
  dplyr::summarize(sessions=sum(sessions), 
                   transactions=sum(transactions), QTY=sum(QTY)) %>%
  as.data.frame()

deviceagg$ECR <- deviceagg$transactions /
  deviceagg$sessions

deviceagg$year_month <- as.yearmon(deviceagg$year_month)

deviceagg <- deviceagg %>% arrange(year_month)

#We can start preparing the data for the second sheet, a month to
#month comparison for the two most recent months. To start, we'll 
#need to create a new aggregate set from deviceagg for each month.
#We can use the same script from earlier but just group by date.

monthlyagg <- monthlysesh %>%
  group_by(year_month) %>%
  dplyr::summarize(sessions=sum(sessions), 
                   transactions=sum(transactions), QTY=sum(QTY)) %>%
  as.data.frame()

monthlyagg

monthlyagg$ECR <- monthlyagg$transactions /
  monthlyagg$sessions

monthlyagg <- monthlyagg %>% mutate(addsToCart = monthlyadds$addsToCart)

#We are only comparing the last two months (both absolute and
#relative differences)

twomonths <- monthlyagg[-(1:10),]

twomonths$year_month <- as.yearmon(twomonths$year_month)

#After creating a data frame with only the last two months, we can
#subtract each second month value with the first month value in order
#to get the difference which will be created in new columns.
#We can then extract this into a new df "absdifferences"

twomonths2 <- twomonths %>% mutate(diffsessions = 
          sessions - lag(sessions), difftransactions = transactions -
          lag(transactions), diffQTY = QTY - lag(QTY), diffECR = ECR - 
          lag(ECR), diffaddsToCart = addsToCart - lag(addsToCart))

absdifferences <- twomonths2[-1,(7:11)]

#We now calculate relative differences by calculating the % change
#from the first month (of the last two)

monthone <- twomonths[1,(2:6)]

reldifferences <- (absdifferences / monthone)*100

library(scales)

reldifferences <- format(reldifferences, format = '%')

#Renaming the columns to avoid confusion.

colnames(reldifferences) <- gsub("^", "rel_", colnames(reldifferences))

#We now have all of our tables needed for the second sheet

install.packages('openxlsx')
library(openxlsx)
library(ggplot2)

wb <- createWorkbook()

addWorksheet(wb, "12 Month Device Analytics")

writeData(wb, "12 Month Device Analytics", deviceagg)

addWorksheet(wb, "Month to Month Comparison")

writeData(wb, twomonths, sheet = "Month to Month Comparison",
               colNames=TRUE)

writeData(wb, absdifferences, sheet = "Month to Month Comparison",
               colNames=TRUE, startRow = 6, startCol = 2)

writeData(wb, reldifferences, sheet = "Month to Month Comparison",
               colNames=TRUE, startRow = 10, startCol = 2)

changesessions <- monthlyagg[(11:12),]
changesessions$ymin <- 0
changesessions$ymax <- 10000000000

ggplot(monthlyagg, aes(x=year_month)) + 
    geom_line(aes(y=sessions, color="sessions")) +
    geom_ribbon(data = changesessions, aes(ymin = sessions + 100000, 
                ymax= sessions - 100000), alpha = 0.2, fill ='green') +
    geom_point(aes(y=sessions, color="sessions"), size = 3,
              shape = 21, fill = 'black') +
    ggtitle("Sessions over 12 Months") +
    scale_color_manual(values = c("sessions" = "red")) +
    scale_x_date(date_breaks = "3 month", date_labels = "%b %Y") +
    labs(x = "Date", y = "Sessions", color = "Variable") +
    theme_classic()

insertPlot(wb, sheet = "Month to Month Comparison", startRow = 2,
           startCol = 8, width = 5, height = 3, units = 'in')

ggplot(monthlyagg, aes(x=year_month)) + 
    geom_line(aes(y=addsToCart, color="addsToCart")) +
    geom_ribbon(data = changesessions, aes(ymin = addsToCart + 10000, 
                ymax= addsToCart - 10000), alpha = 0.2, fill= 'red') +
    geom_point(aes(y=addsToCart, color="addsToCart"), size = 3, 
                shape = 21, fill = 'black') +
    ggtitle("addsToCart over 12 Months") +
    scale_color_manual(values = c("addsToCart" = "blue")) +
    scale_x_date(date_breaks = "3 month", date_labels = "%b %Y") +
    labs(x = "Date", y = "addsToCart", color = "Variable") +
    theme_classic()

insertPlot(wb, sheet = "Month to Month Comparison", startRow = 17,
          startCol = 2, width = 5, height = 3, units = 'in')

ggplot(monthlyagg, aes(x=year_month)) + 
    geom_line(aes(y=transactions, color="transactions")) +
    geom_line(aes(y=QTY, color="QTY")) +
    geom_ribbon(data = changesessions, aes(ymin = transactions + 2000, 
            ymax= transactions - 2000), alpha = 0.2, fill ='green') +
    geom_ribbon(data = changesessions, aes(ymin = QTY + 3500, 
            ymax= QTY - 3500), alpha = 0.2, fill ='green') +
    geom_point(aes(y=transactions, color="transactions"), size = 3, 
              shape = 21, fill = 'black') +
    geom_point(aes(y=QTY, color="QTY"), size = 3, 
            shape = 21, fill = 'black') +
    ggtitle("Transactions/QTY over 12 Months") +
    scale_color_manual(values = c("transactions" = "purple", "QTY" = 'green')) +
    scale_x_date(date_breaks = "3 month", date_labels = "%b %Y") +
    labs(x = "Date", y = "Frequency", color = "Variable") +
    theme_classic()

insertPlot(wb, sheet = "Month to Month Comparison", startRow = 17, 
           startCol = 8, width = 5, height = 3, units = 'in')

ggplot(monthlyagg, aes(x=year_month)) + 
    geom_line(aes(y=ECR, color="ECR")) +
    geom_ribbon(data = changesessions, aes(ymin = ECR + 0.0005, 
              ymax= ECR - 0.0005), alpha = 0.2, fill ='green') +
    geom_point(aes(y=ECR, color="ECR"), size = 3, 
              shape = 21, fill = 'black') +
    ggtitle("ECR over 12 Months") +
    scale_color_manual(values = c("ECR" = "pink")) +
    scale_x_date(date_breaks = "3 month", date_labels = "%b %Y") +
    labs(x = "Date", y = "ECR", color = "Variable") +
    theme_classic()

insertPlot(wb, sheet = "Month to Month Comparison", startRow = 2, startCol = 14, width = 5,
           height = 3, units = 'in')

saveWorkbook(wb, file = "C:/Users/ehaveric99/Documents/ixisexceloutput/12 month analytics summary.xlsx",
             overwrite = TRUE)

#We can now prepare some additional plots to add to our slide deck

deviceagg$year_month <- as.Date(deviceagg$year_month, format = "%m-%Y")

ggplot(deviceagg, aes(x = year_month, y = sessions)) +
  geom_line() +
  geom_point( size = 3, shape = 21, fill = 'white') +
  ggtitle("12 Month Session Data per device used") +
  scale_x_date(date_breaks = "6 month", date_labels = "%b %Y") +
  labs(x = "Month", y = "Sessions") +
  scale_y_continuous(limits = c(0, 600000)) +
  facet_wrap(~ dim_deviceCategory)

ggplot(deviceagg, aes(x = year_month, y = transactions)) +
  geom_line() +
  geom_point( size = 3, shape = 21, fill = 'white') +
  ggtitle("12 Month Transaction Data per device used") +
  scale_x_date(date_breaks = "6 month", date_labels = "%b %Y") +
  labs(x = "Month", y = "Sessions") +
  facet_wrap(~ dim_deviceCategory)

ggplot(deviceagg, aes(x = year_month, y = QTY)) +
  geom_line() +
  geom_point( size = 3, shape = 21, fill = 'white') +
  ggtitle("12 Month QTY purchased Data per device used") +
  scale_x_date(date_breaks = "6 month", date_labels = "%b %Y") +
  labs(x = "Month", y = "Sessions") +
  facet_wrap(~ dim_deviceCategory)

ggplot(deviceagg, aes(x = year_month, y = ECR)) +
  geom_line() +
  geom_point( size = 3, shape = 21, fill = 'white') +
  ggtitle("12 Month ECR Data per device used") +
  scale_x_date(date_breaks = "6 month", date_labels = "%b %Y") +
  labs(x = "Month", y = "Sessions") +
  facet_wrap(~ dim_deviceCategory)

