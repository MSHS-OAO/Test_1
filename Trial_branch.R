# Get today and yesterday's date

Today <- as.timeDate(format(Sys.Date(),"%m/%d/%Y"))
#Today <- as.timeDate(as.Date("07/29/2020", format = "%m/%d/%Y"))


#Determine if yesterday was a holiday/weekend 
#get yesterday's DOW
Yesterday <- as.timeDate(format(Sys.Date()-1,"%m/%d/%Y"))
#Yesterday <- as.timeDate(as.Date("07/28/2020", format = "%m/%d/%Y"))

#Get yesterday's DOW
Yesterday_Day <- dayOfWeek(Yesterday)

#Remove Good Friday from MSHS Holidays
NYSE_Holidays <- as.Date(holidayNYSE(year = 1990:2100))
GoodFriday <- as.Date(GoodFriday())
MSHS_Holiday <- NYSE_Holidays[GoodFriday != NYSE_Holidays]

#Determine whether yesterday was a holiday/weekend
Holiday_Det <- isHoliday(Yesterday, holidays = MSHS_Holiday)

#Set up a default calendar for collect to received TAT calculations for Pathology and Cytology
create.calendar("MSHS_working_days", MSHS_Holiday, weekdays=c("saturday","sunday"))
bizdays.options$set(default.calendar="MSHS_working_days")


# Select file/folder path for easier file selection and navigation
# user_wd <- choose.dir(caption = "Select your working directory")
#user_wd <- "J:\\deans\\Presidents\\HSPI-PM\\Operations Analytics and Optimization\\Projects\\Service Lines\\Lab KPI\\Data"
#user_path <- paste0(user_wd, "\\*.*")

if (list.files("J://") == "Presidents") {
  user_directory <- "J:/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/Service Lines/Lab Kpi/Data"
} else {
  user_directory <- "J:/deans/Presidents/HSPI-PM/Operations Analytics and Optimization/Projects/Service Lines/Lab Kpi/Data"
}
user_path <- paste0(user_directory, "/*.*")

