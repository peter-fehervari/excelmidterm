########### Libraries #############
library(xlsx)
#library(openxlsx)
########### Description #############

# The dataset is dogs with babesiosis tests. All values are randomized for each iteration.
# Data: 
  # ID: 1:n values
  # name: randomized from list (Dog_names.csv)
  # sex: randomized 
  # age: random value in months, range 10-156
  # body mass: random from normal dist, with different mu and sigma for the sexes
  # protected against ticks?
  # tick confirmed
  # body temp: in celsius 38-39.2 over 39.2 is fever (upto 41)  
  # diagnosis blood smear test: random, with increasing probability with no tick, tick confirmed and fever

########## Init #####################
#set.seed(1717)
tests <- 3 # number of total tests to generate
Dognames <- read.table("Dog_names.csv",h = T ,sep = ",")
N <- 400 # Number of total dogs 
Prob_male <- 0.5 # Sex ratio  
Min_age_mon <- 10 #in months, not too small to avoid simulating juv. dog weights
Max_age_mon <- 156
Mu_body_mass_male <- 16
Mu_body_mass_female <- 11
Sigma_body_mass_male <- 4
Sigma_body_mass_female <- 3
Tick_protection_prevalence <- 0.3 # Tick protection prevalence for all dogs r
Tick_confirmation_prob <- 0.5 # P(dog had a confirmed tick| no protection) 
Temp_min <- 38
Temp_max <- 41 #Fever from 39.2

# quick logit to prob converter
logit2prob <- function(logit){
  odds <- exp(logit)
  prob <- odds / (1 + odds) 
  return(prob)
}
# level selector for randomly chosen vars. 
level_selector<-function(x){sample(levels(data[,x]),1)}


########## Generate Midterms ############
for ( i in 1:tests)
{
  
id <- 1:N 

dognames <-ifelse(sex=="Male", yes = sample(Dognames$MALE.DOG.NAMES, replace = F),
                  no = sample(Dognames$FEMALE.DOG.NAMES, replace = F))

sex <- sample(c("Male","Female"), size = N, prob = c(Prob_male, 1-Prob_male),replace = T)

age <-sample(Min_age_mon:Max_age_mon, size=N, replace = T)

body_mass <- round(ifelse(sex=="Male", yes = rnorm(N,Mu_body_mass_male,Sigma_body_mass_male) ,
                  no = rnorm(N,Mu_body_mass_female,Sigma_body_mass_female)),1)

tick_protection <- sample(c("Yes","No"), size = N, 
              prob = c(Tick_protection_prevalence, 1-Tick_protection_prevalence ),replace = T)

tick_confirmed <- ifelse(tick_protection == "Yes", 
                         yes = sample(c("Yes","No"),prob = c(Tick_confirmation_prob,1-Tick_confirmation_prob)),
                         no = "No")

body_temp <- round(runif(N,min = Temp_min,max = Temp_max),1)

#Create near final dataframe
data <- data.frame(id=id, dognames=factor(dognames), sex=factor(sex), body_mass=body_mass, age=age,
                   tick_protection = factor(tick_protection), tick_confirmed=factor(tick_confirmed),
                   body_temp = body_temp)

#Smear diagnosis  will depend on tick confirmed presence and fever presence.
#Calculating fever in excel using IF will be a task, thus it will be excluded from the table

fever <- factor(ifelse(body_temp>39.2,"Fever","Normal"))

#logit_dis <- 0.9*(as.numeric(data$tick_confirmed)-1)+0.02*data$body_temp+rnorm(N,0,1)
logit_dis <- 1*(as.numeric(data$tick_confirmed)-1)+-0.5*as.numeric(fever)-1+rnorm(N,0,1)

data$smear_positive <- factor(ifelse(logit2prob(logit_dis)>0.5,"Positive","Negative"))

########## Creating individual tasks #############
# General inits
id_vars <- c("dognames","id")
colors_sel <- c("black","red","green","yellow","blue","pink")
line_type <- c("solid","dashed","dotted")
factor_vars <-c("sex","tick_protection","tick_confirmed","smear_positive")
numeric_vars <-c("body_mass","body_temp","age")


### Formatting tasks
#inits
format_alignment_var <- sample(id_vars,1)
format_color_var <- sample(colors_sel,1)
format_line_type_var <- sample (line_type,1)

tasks_table_formatting <- data.frame(Formatting= c(" Increase the font size to 14pts and set font style to Bold for coloumn headers (first row).",
                                                   "Adjust coloumn width to fit data for every coloumn and align coloumn headers to center.",
                                                   paste("Align all values in the",format_alignment_var,"column to center."),
                                                   paste("Create a",format_color_var ,"border around the coloumn headers using thick", format_line_type_var,"lines."),
                                                   "Freeze the top row,",
                                                   "Sort the data ascending by age."))

#### Function tasks
#inits
countif_var <- c(sample(factor_vars,1))
countif_var_level <- level_selector(countif_var)

averageif_var <- sample(numeric_vars,1)
averageif_condition_var <- c(sample(factor_vars,1))
averageif_condition_var_level <- level_selector(averageif_condition_var)


tasks_table_functions <- data.frame(Functions = c("Type Fever in cell J1. For each row, evaluate whether the body_temp is higher than 39.1 C. Type YES if the value is over and NO if it is not.",
                                                  paste("In cell L2, calculate the number of individuals in the" ,countif_var , "coloumn that have the value of",countif_var_level),
                                                  paste("In cell L3, calculate the mean",averageif_var,"for observations where",averageif_condition_var,"is",averageif_condition_var_level )
                                                 ))

####Chart tasks
#inits

boxplot_cat_var <- c(sample(factor_vars,1))
boxplot_num_var <- sample(numeric_vars,1)

scatterplot_num_var <- sample(numeric_vars,1)
scatterplot_num_var2 <- sample(numeric_vars,1)

tasks_table_charts <- data.frame(Charts =c(paste("Create a boxplot of the", boxplot_num_var, "variable. Use the",boxplot_cat_var,"variable for grouping."),
                                           "Add an appropriate title for the chart. Change the colours of the boxes to black and white. Add appropriate Y axis name.",
                                           paste("Create a scatter chart with", scatterplot_num_var, "and", scatterplot_num_var2, "variables."),
                                           "Add an appropriate title for the chart.Change the dot colours to black. Add appropriate X and Y axis names. Delete horizontal and vertical grid lines."
                                           ))

#### Pivot tasks
#inits
pivot_cat_var1 <- c(sample(factor_vars,1))
pivot_cat_var2 <- c(sample(factor_vars,1))
pivot_num_var <- sample(numeric_vars,1)

tasks_table_pivot <- data.frame(Pivot=c(paste("Create a pivot table on a new sheet. For row variables select", pivot_cat_var1, "for coloumns, select the",pivot_cat_var2,"variable."),
                                        paste("For values, use the",pivot_num_var, "variable. Change the value field settings of", pivot_num_var, "to Average."),
                                        "Select the most distasteful design (from the predefined list) for the pivot table.",
                                        "Insert a pivot coloumn chart.",
                                        "Add an appropriate title for the chart. Add appropriate Y axis label. Change the coloumn colors to black and white."
                                              ))

#Points
tasks_table_points <- data.frame(Points=c(1,1,1,1,2,2,"","",2,2,2,"","",4,2,4,2,"","",4,2,2,4,2))

######## Creating the workbook ##############
wb<-createWorkbook(type="xlsx")

# Data sheet 
sheet <- createSheet(wb, sheetName = "Data")
addDataFrame(data, sheet, startRow=1, startColumn=1,row.names = FALSE)

# Tasks sheet
sheet2 <- createSheet(wb, sheetName = "Tasks")
cell_style_task_title <- CellStyle(wb) + 
  Font(wb, heightInPoints=12, isBold=TRUE) +
  Alignment(h="ALIGN_CENTER")

cell_style_task_points <- CellStyle(wb) + 
  Font(wb) +
  Alignment(h="ALIGN_CENTER")

addDataFrame(tasks_table_formatting, sheet2, startRow=2, startColumn=3,row.names = FALSE,colnamesStyle = cell_style_task_title)
addDataFrame(tasks_table_functions, sheet2, startRow=10, startColumn=3,row.names = FALSE,colnamesStyle = cell_style_task_title)
addDataFrame(tasks_table_charts, sheet2, startRow=15, startColumn=3,row.names = FALSE,colnamesStyle = cell_style_task_title)
addDataFrame(tasks_table_pivot, sheet2, startRow=21, startColumn=3,row.names = FALSE,colnamesStyle = cell_style_task_title)
addDataFrame(tasks_table_points, sheet2, startRow=2, startColumn=2,row.names = FALSE,colnamesStyle = cell_style_task_title,)

autoSizeColumn(sheet2, colIndex=3)
# Export
saveWorkbook(wb, paste0("midterm",i,".xlsx"))
}
