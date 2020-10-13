########### Libraries #############
library(xlsx)

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
  # body temp: in celsius 38-39.2, fever upto 41  
  # diagnosis blood smear test: random, with increasing probability with no tick, tick confirmed and fever

########## Init #####################
set.seed(1717)

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

logit2prob <- function(logit){
  odds <- exp(logit)
  prob <- odds / (1 + odds) # quick logitto prob concverter
  return(prob)
}
########## Generate data ############

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
level_selector<-function(x){sample(levels(data[,x]),1)}

# Formatting tasks
tasks_table_formatting <- data.frame(Formatting= c(" Increase the font size to 14pts and set font style to Bold for coloumn headers (first row).",
                                   "Adjust coloumn width to fit data for every coloumn and align coloumn headers to center.",
                                   paste("Align all values in the",sample(id_vars,1),"column to center."),
                                   paste("Create a",sample(colors_sel,1) ,"border around the coloumn headers using thick", sample (line_type,1),"lines."),
                                   "Freeze the top row,"))

# Function tasks
#inits
countif_var <- c(sample(factor_vars,1))
countif_var_level <- level_selector(countif_var)

tasks_table_functions <- data.frame(Functions = c(paste("Calculate the number of individuals in the" ,countif_var , "coloumn that have the value of",countif_var_level ,". Use functions!")
                                                 ))
        


wb<-createWorkbook(type="xlsx")
sheet <- createSheet(wb, sheetName = "Data")
addDataFrame(data, sheet, startRow=1, startColumn=1,row.names = FALSE)
addDataFrame(tasks, sheet, startRow=1, startColumn=12,row.names = FALSE)
addDataFrame(tasks2, sheet, startRow=6, startColumn=12,row.names = FALSE)
saveWorkbook(wb, "Data_trial.xlsx")
