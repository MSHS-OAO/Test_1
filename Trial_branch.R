# Get today and yesterday's date
#### trial git fetch vs git pull
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

if(((Holiday_Det) & (Yesterday_Day =="Mon"))|((Yesterday_Day =="Sun") & (isHoliday(Yesterday-(86400*2))))){ # Scenario 1: Mon Holiday or Friday Holiday (Need to select 4 files)
  # Save scenario
  scenario <- 1
  # Import SCC data
  SCC_Holiday_Monday_or_Friday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Recent Holiday"), sheet = 1, col_names = TRUE)
  SCC_Sunday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Sunday"), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Saturday"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(SCC_Holiday_Monday_or_Friday, SCC_Sunday, SCC_Saturday)
  SCC_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE)
  # 
  # Import Sunquest data
  SQ_Holiday_Monday_or_Friday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Recent Holiday"), sheet = 1, col_names = TRUE))
  SQ_Sunday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Sunday"), sheet = 1, col_names = TRUE))
  SQ_Saturday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Saturday"), sheet = 1, col_names = TRUE))
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(SQ_Holiday_Monday_or_Friday, SQ_Sunday, SQ_Saturday)
  SQ_Weekday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE))
  #
  # Import Powerpath data
  PP_Holiday_Monday_or_Friday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Recent Holiday"), skip = 1, 1)
  PP_Holiday_Monday_or_Friday  <- PP_Holiday_Monday_or_Friday[-nrow(PP_Holiday_Monday_or_Friday),]
  PP_Sunday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Sunday"), skip = 1, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Saturday"), skip = 1, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- data.frame(rbind(PP_Holiday_Monday_or_Friday ,PP_Sunday,PP_Saturday),stringsAsFactors = FALSE)
  PP_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Most Recent Weekday"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),],stringsAsFactors = FALSE)
} else if ((Holiday_Det) & (Yesterday_Day =="Sun")){ # Scenario 2: Regular Monday (Need to select 3 files)
  # Save scenario
  scenario <- 2
  #
  # Import SCC data
  SCC_Sunday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Sunday"), sheet = 1, col_names = TRUE)
  SCC_Saturday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Saturday"), sheet = 1, col_names = TRUE)
  #Merge the weekend data with the holiday data in one data frame
  SCC_Not_Weekday <- rbind(SCC_Sunday,SCC_Saturday)
  SCC_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE)
  #
  # Import Sunquest data
  SQ_Sunday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Sunday"), sheet = 1, col_names = TRUE))
  SQ_Saturday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Saturday"), sheet = 1, col_names = TRUE))
  #Merge the weekend data with the holiday data in one data frame
  SQ_Not_Weekday <- rbind(SQ_Sunday,SQ_Saturday)
  SQ_Weekday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE))
  #
  # Import Powerpath data
  PP_Sunday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Sunday"), skip = 1, 1)
  PP_Sunday <- PP_Sunday[-nrow(PP_Sunday),]
  PP_Saturday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Saturday"), skip = 1, 1)
  PP_Saturday <- PP_Saturday[-nrow(PP_Saturday),]
  #Merge the weekend data with the holiday data in one data frame
  PP_Not_Weekday <- data.frame(rbind(PP_Sunday,PP_Saturday),stringsAsFactors = FALSE)
  PP_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Most Recent Weekday"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
} else if ((Holiday_Det) & ((Yesterday_Day !="Mon")|(Yesterday_Day !="Sun"))){ #Scenario 3: Midweek holiday (Need to select 2 files)
  # Save scenario
  scenario <- 3
  #
  # Import SCC data
  SCC_Holiday_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Recent Holiday"), sheet = 1, col_names = TRUE)
  SCC_Not_Weekday <- SCC_Holiday_Weekday
  SCC_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE)
  #
  # Import Sunquest data
  SQ_Holiday_Weekday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Recent Holiday"), sheet = 1, col_names = TRUE))
  SQ_Not_Weekday <- SQ_Holiday_Weekday
  SQ_Weekday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select Sunquest Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE))
  #
  # Import Powerpath data
  PP_Holiday_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Recent Holiday"), skip = 1, 1)
  PP_Holiday_Weekday <- PP_Holiday_Weekday[-nrow(PP_Holiday_Weekday),]
  PP_Not_Weekday <- data.frame(PP_Holiday_Weekday, stringsAsFactors = FALSE)
  PP_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Most Recent Weekday"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
} else { #Scenario 4: Tue-Fri without holidays (Need to select 1 file)
  # Save scenario
  scenario <- 4
  #
  # Import SCC data
  SCC_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/SCC CP Reports/*.*"), caption = "Select SCC Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE)
  SCC_Not_Weekday <- NULL
  #
  # Import Sunquest data
  SQ_Weekday <- suppressWarnings(read_excel(choose.files(default = paste0(user_directory, "/SUN CP Reports/*.*"), caption = "Select SunQuest Report for Labs Resulted on Most Recent Weekday"), sheet = 1, col_names = TRUE))
  SQ_Not_Weekday <- NULL
  #
  # Import Powerpath data
  PP_Weekday <- read_excel(choose.files(default = paste0(user_directory, "/AP & Cytology Signed Cases Reports/*.*"), caption = "Select PowerPath Report for Specimens Signed Out on Most Recent Weekday"), skip = 1, 1)
  PP_Weekday <- data.frame(PP_Weekday[-nrow(PP_Weekday),], stringsAsFactors = FALSE)
  PP_Not_Weekday <- NULL
}







