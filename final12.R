#the necessary libraries used for this project
library(tidyverse)
library(readxl)
library(lubridate)
library(stringr)
library(ggplot2)
library(plotly)
library(anytime)
library(stats)
library(writexl)
library(car)
library(factoextra)
library(GGally)
library(gridExtra)
library(ggthemes)


#import the data sets
df_12 <- read_excel("C:/Users/nebar/Desktop/project/data2/data12.xlsx")

dat3_12 <- read_excel("C:/Users/nebar/Desktop/project/data3/UPC12.xlsx")

class(df_12)

#Remove some of the variables
df3 <- dat3_12|>
  select(-c(2,3,4,5,6,12,18,22))

df3_a <- df3|>
  mutate(x  = Connections - df3$`Total Retailer Connections` )

#rename the variables
df <- df3_a |> 
  rename(Total_connections = Connections,  Missed_Connection = x, 
         ORG_ID = `ORG ID`,
         First_Subscription_Date = `First Subscription Date`,
         Doc_Flow_Total_Monthly_Average = `Doc Flow Total Monthly Average`,
         Industry_SPS = `Industry SPS`,Last_Updated = Updated,
         Industry_Segment = `Industry Segment`,
         Industry_Vertical = `Industry Vertical`,
         Total_Retailer_Connections = `Total Retailer Connections`)|>
  select(EAN, Title, Merchant,Last_Updated,CONSUMERPACKAGECODE,ORG_ID,
         First_Subscription_Date, Doc_Flow_Total_Monthly_Average,
         Industry_SPS,AnnualRevenue,Industry_Segment,Industry_Vertical,
         Total_Retailer_Connections,Total_connections,Missed_Connection)


#change the df_12 to data frame and rename some of the variables
df12_mod <- data.frame(df_12)|>
  rename(Last_Updated = last_updated,  Missed_Connection =  Missed_connections)

#merge the two data sets
df_final <- bind_rows(df,df12_mod)

summary(df_final)

#change some of the variables from a character to factor for analysis
df_final$Industry_Segment <- as.factor(df_final$Industry_Segment)

df_final$Industry_SPS <- as.factor(df_final$Industry_SPS)

df_final$Industry_Vertical <- as.factor(df_final$Industry_Vertical)

#change first subscription date from char to date
df_final$First_Subscription_Date <- anydate(df_final$First_Subscription_Date)


#drop the null values for doc but replace 0 for annual revenue with na
final12u <- df_final|>
  drop_na(Doc_Flow_Total_Monthly_Average)|>
  mutate(AnnualRevenue = ifelse(is.na(AnnualRevenue), 0, AnnualRevenue))

#change the dataset to a dataframe
final12u <- data.frame(final12u)

#drop any null values and duplicate rows
final12ua <- final12u|>
  drop_na()|>
  distinct(EAN,.keep_all = TRUE)

#top 50 different companies with higher missed connections
top <- final12ua |>
  arrange(desc(Missed_Connection),.by_group = TRUE)|>
  slice(1:102)

#calculate the count and mean of some of the variables by industry vertical
top50_Ver <- top|>
  group_by(Industry_Vertical)|>
  summarize(count = length(Industry_Vertical),
          avg_rev = mean(AnnualRevenue),
          avg_ret_con = mean(Total_Retailer_Connections),
          avg_mis_con = mean(Missed_Connection))

#read customer data set
cus_df <- read_excel('C:/Users/nebar/Desktop/project/customer_details1.xlsx')

#rename org id to ease merge
cus_df1 <- cus_df|>
  rename(ORG_ID = `ORG ID`)
company_id <- inner_join(cus_df1,top, by = "ORG_ID")

#keep the necessary variables only
top_50 <- company_id|>
  select(-c(2, 4:16),)

#export the top50 as excel file
write_xlsx(top_50, 'C:/Users/nebar/Desktop/project/data2/top_50.xlsx')

#records with docflow != 0 but total_retailer equal to 0
con0 <- final12ua|>
  filter(Doc_Flow_Total_Monthly_Average != 0 & Total_Retailer_Connections == 0)

#export the con0 as excel file
write_xlsx(con0, 'C:/Users/nebar/Desktop/project/data2/con0.xlsx')

#plot industry vertical for the data frame con0
ggplotly(ggplot(data = con0, aes(x = Industry_Vertical)) + geom_bar() + 
          labs(title = "Industries with no Connection")+
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#plot Total Connections by Industry Vertical
ggplotly(ggplot(data = final12ua, aes(x = Industry_Vertical, y = Total_connections)) + 
          geom_boxplot() + labs(title = "Total Connections by Industry Vertical")+
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))


#plot Annualrevenue by Industry SPS
ggplotly(ggplot(data = final12ua, aes(x = Industry_SPS, y = AnnualRevenue)) + 
          geom_boxplot() + labs(title = "AnnualRevenue by Industry SPS") +
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#plot MissedConnection by Industry SPS
ggplotly(ggplot(data = final12ua, aes(x = Industry_SPS, y = Missed_Connection)) + 
          geom_boxplot() + labs(title = "MissedConnection by Industry SPS") +
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#plot annual revenue by industry vertical
ggplotly(ggplot(data = final12ua, aes(x = Industry_Vertical, y = AnnualRevenue)) + 
          geom_boxplot() + 
          labs(title = "Annual Revenue by Industry Vertical/UPC12", 
            x = "Industry Vertical",y = "Annual Revenue") + 
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#plot Missed connection by industry vertical
ggplotly(ggplot(data = final12ua, aes(x = Industry_Vertical, y = Missed_Connection)) + 
          geom_boxplot() + 
          labs(title = "Missed Connection by Industry Vertical", x = "Industry Vertical",
          y = "Missed Connection") +
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#distribution of industry vertical for the data frame
ggplotly(ggplot(data = final12ua, aes(x = Industry_Vertical)) + geom_bar() + 
          labs(title = "Industries Distribution for UPC12")+
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#Records with missed connection is 0/companies using only SPS
miss_no <- final12ua|>
  filter(Missed_Connection == 0)

#plot industry vertical distribution for those with no missed connection
ggplotly(ggplot(data = miss_no, aes(x = Industry_Vertical)) + geom_bar() + 
          labs(title = "Industries with No Missed Connection")+
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))


#export the outputs to excel file
write_xlsx(final12ua, 'C:/Users/nebar/Desktop/project/data2/final12u.xlsx')

#records with missed connection less than 0
less <- final12ua|>
  filter(Missed_Connection < 0)

#plot industry vertical distribution for those with less connection 
ggplotly(ggplot(data = less, aes(x = Industry_Vertical)) + geom_bar() + 
          labs(title = "Industries with Missed Connection Less Than 0/UPC12")+
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#records with missed connection greater than 0
more <- final12ua|>
  filter(Missed_Connection > 0)

#plot industry vertical distribution for those with more connection
ggplotly(ggplot(data = more, aes(x = Industry_Vertical)) + geom_bar() +
          labs(title = "Industries with Missed Connection Greater Than 0/UPC12")+
          theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#Assigning values as less, more and equal in a new column
final12ua$Miss <- ifelse(final12ua$Missed_Connection < 0, "less",
                 ifelse(final12ua$Missed_Connection > 0, "more", "equal"))
final12ua$Miss <- as.factor(final12ua$Miss)

#correlation among numerical variables
ggpairs(data = final12ua,columns = c(8,10,13,14,15))

#Multiple Regression
m <- lm(Missed_Connection ~ Industry_Vertical + Doc_Flow_Total_Monthly_Average +
          Total_connections + Total_Retailer_Connections, data = final12ua)
summary(m)

#Detecting outliers and put the three graphs in a one row
par(mfrow = c(1,3))

hist(final12ua$Missed_Connection,main = "histogram")

boxplot(final12ua$Missed_Connection,main = "boxplot")

qqnorm(final12ua$Missed_Connection,main = "Normal Q-Q plot")


#removing the outliers
quartiles <- quantile(final12ua$Missed_Connection, probs=c(.25, .75), na.rm = FALSE)

IQR <- IQR(final12ua$Missed_Connection)

Lower <- quartiles[1] - 1.5*IQR

Upper <- quartiles[2] + 1.5*IQR 

finall12_no  <- subset(final12ua, final12ua$Missed_Connection > Lower & 
                        final12ua$Missed_Connection < Upper)

hist(finall12_no$Missed_Connection,main = "histogram")

boxplot(finall12_no$Missed_Connection,main = "boxplot")

qqnorm(finall12_no$Missed_Connection,main = "Normal Q-Q plot")

#Multiple regression and correlation for dataset without the outliers
mo <- lm(Missed_Connection ~ Industry_Vertical + AnnualRevenue + 
          Doc_Flow_Total_Monthly_Average + 
          Total_Retailer_Connections, data = finall12_no)
summary(mo)

#correlation without outliers
ggpairs(data = finall12_no,columns = c(8,10,13,14,15))


#Distribution of variables
ggplotly(ggplot(data = final12ua, aes(x = Missed_Connection))+ geom_bar() + 
          labs(title = "Distribution of Missed Connection"))

ggplotly(ggplot(data = final12ua, aes(x = Total_Retailer_Connections)) + 
          geom_bar() + labs(title = "Distribution of Total Retailer Connection"))

ggplotly(ggplot(data = final12ua, aes(x = AnnualRevenue))+ geom_histogram(bins = 30) +
          labs(title = "Distribution of Annual Revenue"))

#relationship between annual revenue and total connections by industry vertical
ggplotly(ggplot(data = final12ua, 
              aes(x = Missed_Connection, y = AnnualRevenue,color = Industry_Vertical)) + 
              geom_point() + 
              labs(title = "Relationship of Annual Revenue and Missed Connection 
                  by Industry Vertical/UPC12"))

#One-way ANOVA
#H0: mu(less) = mu(equal) = mu(more)   Ha: at least two of the means are different
mar1 <- aov(AnnualRevenue ~ Miss, data = final12ua)

summary(mar1)
#pvalue < 0 .05 , Reject H0. Therefore, there is sufficient evidence to support 
#the claim that at least two of the average revenues are different

TukeyHSD(mar1)
# There is a significant difference in revenue among the three groups

#error bar among the three groups
final12ua|>
  group_by(Miss) |>
  summarize(avere = mean(AnnualRevenue),
            sere = sd(AnnualRevenue)/sqrt(length(AnnualRevenue)),
            tstar=qt(1-0.05/2,length(AnnualRevenue)-1))|>
  ggplot(aes(x = Miss,y=avere))+ geom_point() + 
  geom_errorbar(aes(ymin = avere-tstar*sere,ymax=avere + tstar*sere))

#take sample data for clustering purpose using stratified sampling 
set.seed(1)

sample <- finall12_no|>  # used the dataset without outliers
  group_by(Industry_Vertical)|>
  sample_frac(0.01)  


#cluster Analysis(kmeans)
final_12d <- scale(sample[, c("AnnualRevenue","Doc_Flow_Total_Monthly_Average",
                              "Total_Retailer_Connections","Missed_Connection")])

fviz_nbclust(final_12d,kmeans,method = "wss")

set.seed(100)

kmodeld <- kmeans(final_12d,centers = 5,nstart = 50)

kmodeld

fviz_cluster(kmodeld, data = final_12d)

round(kmodeld$centers,2)

#AnnualRevenue trend overtime
ggplotly(ggplot(data = final12ua, aes(x = First_Subscription_Date,y = AnnualRevenue))+ 
          geom_line() + labs(title = "Annual Revenue Trend over Time"))

#total connections trend overtime
ggplotly(ggplot(data = final12ua, aes(x = First_Subscription_Date,y = Total_connections))+ 
          geom_line(color = "tomato3") + labs(title = "Total Connection over Time"))


#preprocess the lastupdate column
final12ua <- final12ua |>
  mutate(
    # Remove brackets and quotes, then split into lists
    Last_Updated = gsub("\\[|\\]|'", "", Last_Updated),
    Last_Updated = strsplit(Last_Updated, ", "),
    First_Subscription_Date = mdy(First_Subscription_Date))

#disaggregate the lastupdated column into individual rows
final12 <- final12ua |>
  unnest(Last_Updated) |>
  mutate(Last_Updated = ymd(Last_Updated)) 

#filter out those records which have missed connection with in the last 5years
start_date <- as.Date("2020-01-01")

last5 <- final12|>
  group_by(EAN)|>
  filter(Last_Updated >= start_date)|>
  mutate(count_5year = length(Last_Updated))

#calculate the count and mean of connection by industry vertical
last5_sum <- last5|>
  group_by(Industry_Vertical, EAN)|>
  summarize(count = unique(count_5year))

#plot the count of total connection of last5 dataframe by Industry vertical
ggplotly(ggplot(data = last5_sum, aes(x = Industry_Vertical, y = count))+ 
  geom_boxplot() +  
  labs(title = "Total Connection (2020-01-01 to 2024-10-31) by Industry Vertical",
       x = "Industry Vertical",y = "count") +
  theme(axis.text.x = element_text(angle = 90, hjust = 1)))

#Compare lastupdated and First_Subscription_Date
final <- final12 |>
  mutate(
    Is_After = Last_Updated > First_Subscription_Date,  
    Is_Before = Last_Updated < First_Subscription_Date 
  )

#records which have used other connections after they joined sps
after <- final|>
  filter(Is_After == TRUE)

#records which have used other connections before they joined sps
before <- final|>
  filter(Is_Before == TRUE)

#after datasets by industry vertical
ggplot(data = after, aes(x = Industry_Vertical))+ geom_bar() +
  labs(title = "Industries with Connection After Subscription")+
  theme(axis.text.x = element_text(angle = 90, hjust = 1))

#before datasets by industry vertical
ggplot(data = before, aes(x = Industry_Vertical))+ geom_bar() + 
  labs(title = "Industries with Connection Before Subscription")+
  theme(axis.text.x = element_text(angle = 90, hjust = 1))
