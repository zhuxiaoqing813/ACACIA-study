# 0 安装并加载包
library(randomForest)
library(MatchIt)
library(ggplot2)
library(dplyr)
library(readr)

#help(MatchIt)
# 1 输入数据
## 1.1 读取匹配前数据
#data0 <- read.csv("D:/ACACIA 分析/分析数据/ACACIA_new/250523/Acacia原始数据250708-1200.csv")
data0 <- read.csv("D:/ACACIA 分析/分析数据/250708/Acacia原始数据250710-2300.csv")
### 定义变量类型
#### 因子变量1
data0$case <- as.factor(data0$case)
data0$Area <- as.factor(data0$Area)
data0$Province <- as.factor(data0$Province)
data0$Registration <- as.factor(data0$Registration)
data0$Ownership <- as.factor(data0$Ownership)
data0$Hospital.level <- as.factor(data0$Hospital.level)
data0$Institution.type <- as.factor(data0$Institution.type)
data0$Management.type <- as.factor(data0$Management.type)
data0$Doctor.title <- as.factor(data0$Doctor.title)
data0$Doctor.certificate <- as.factor(data0$Doctor.certificate)
data0$Doctor.age <- as.factor(data0$Doctor.age)
data0$Doctor.education <- as.factor(data0$Doctor.education)
data0$Patient.gender <- as.factor(data0$Patient.gender)
data0$Doctor.ethnicity <- as.factor(data0$Doctor.ethnicity)
#### 因子变量2
data0$Diagnosis.quality <- as.factor(data0$Diagnosis.quality)
data0$Treatment_quality.surgical <- as.factor(data0$Treatment_quality.surgical)
data0$Treatment_quality.western_medicine <- as.factor(data0$Treatment_quality.western_medicine)
data0$Treatment_quality.traditional_Chinese_medicine<- as.factor(data0$Treatment_quality.traditional_Chinese_medicine)
data0$Treatment_quality.physical_therapy<- as.factor(data0$Treatment_quality.physical_therapy)
data0$Treatment_quality.psychological_therapy<- as.factor(data0$Treatment_quality.health_education_and_lifestyle_intervention)
data0$Treatment_quality.physical_therapy<- as.factor(data0$Treatment_quality.physical_therapy)
data0$Treatment_quality.referral<- as.factor(data0$Treatment_quality.referral)
data0$Treatment_quality.followup<- as.factor(data0$Treatment_quality.followup)
data0$opening.time<- as.factor(data0$opening.time)
data0$Pharmaceutical_services.Satisfaction<- as.factor(data0$Pharmaceutical_services.Satisfaction)
#### 数值变量1
data0$Treatment.score <- as.numeric(data0$Treatment.score)
#### 数值变量2
data0$Consultations.proportion <- as.numeric(data0$Consultations.proportion)
data0$Physical_examinations.proportion <- as.numeric(data0$Physical_examinations.proportion)
data0$Laboratory_test.proportion <- as.numeric(data0$Laboratory_test.proportion)
data0$Inappropriate_tests.number <- as.numeric(data0$Inappropriate_tests.number)
#### 数值变量3
data0$Inappropriate_surgical_treatments.number <- as.numeric(data0$Inappropriate_surgical_treatments.number)
data0$Inappropriate_Westernmedicine_treatments.number<- as.numeric(data0$Inappropriate_Westernmedicine_treatments.number)
data0$Inappropriate_surgical_treatments.number_A<- as.numeric(data0$Inappropriate_surgical_treatments.number_A)
data0$Inappropriate_surgical_treatments.number_B<- as.numeric(data0$Inappropriate_surgical_treatments.number_B)
data0$Inappropriate_surgical_treatments.number_C<- as.numeric(data0$Inappropriate_surgical_treatments.number_C)
data0$Inappropriate_surgical_treatments.number_D<- as.numeric(data0$Inappropriate_surgical_treatments.number_D)
data0$Inappropriate_surgical_treatments.number_E<- as.numeric(data0$Inappropriate_surgical_treatments.number_E)
data0$Inappropriate_surgical_treatments.number_F<- as.numeric(data0$Inappropriate_surgical_treatments.number_F)
#### 数值变量4
data0$PPPC.score <- as.numeric(data0$PPPC.score)
data0$PPPC_Disease_and_illness_experience <- as.numeric(data0$PPPC_Disease_and_illness_experience)
data0$PPPC_Common_ground <- as.numeric(data0$PPPC_Common_ground)
data0$PPPC_Whole_person <- as.numeric(data0$PPPC_Whole_person)
data0$PPPC_Patient_doctor_relationship <- as.numeric(data0$PPPC_Patient_doctor_relationship)
#### 数值变量5
data0$PPPC.score <- as.numeric(data0$PPPC.score)
data0$PPPC_Disease_and_illness_experience <- as.numeric(data0$PPPC_Disease_and_illness_experience)
data0$PPPC_Common_ground <- as.numeric(data0$PPPC_Common_ground)
data0$PPPC_Whole_person <- as.numeric(data0$PPPC_Whole_person)
data0$PPPC_Patient_doctor_relationship <- as.numeric(data0$PPPC_Patient_doctor_relationship)
#### 数值变量6
data0$Consultation.cost <- as.numeric(data0$Consultation.cost)
data0$Examination.cost <- as.numeric(data0$Examination.cost)
data0$Treatment.cost <- as.numeric(data0$Treatment.cost)
data0$Other.cost <- as.numeric(data0$Other.cost)
data0$Medicine.cost <- as.numeric(data0$Medicine.cost)
data0$Total.cost <- as.numeric(data0$Total.cost)
data0$Waiting_time <- as.numeric(data0$Waiting_time)
data0$Pharmacy_processing_time <- as.numeric(data0$Pharmacy_processing_time)
data0$Consultation_room.time <- as.numeric(data0$Consultation_room.time)
data0$Total.time <- as.numeric(data0$Total.time)

## 检查 Ownership 列的缺失值数量
sum(is.na(data0$Ownership))
## 删除 Ownership 列中存在缺失值的行
data_clean <- data0[!is.na(data0$Ownership), ]
## 验证缺失值是否已删除
sum(is.na(data_clean$Ownership))  # 应为0

## 1.2.4 平衡性检验
library(cobalt)
bal.tab(Ownership ~ Area, data = data_clean)
bal.tab(Ownership ~ Management.type, data = data_clean)

# 2 填补协变量缺失数据
# 2.1 查看各列缺失值数量
colSums(is.na(data_clean)) 
## 可视化缺失模式
library(visdat)
vis_miss(data_clean)
## 找出包含缺失值的完整行
na_rows <- data_clean[!complete.cases(data_clean), ]
head(na_rows)

data_imputed <- data_clean

data_imputed$case <- factor(data_imputed$case, levels = c("1", "2", "3", "4", "6", "7" ,"8", "9", "10", "11", "12"),
                            labels = c("Migraine", 
                                       "Postpartum depression",
                                       "Pediatric diarrhea", 
                                       "Common cold",
                                       "Gastritis", 
                                       "Asthma",
                                       "Angina pectoris",
                                       "Back pain",
                                       "Stress urinary incontinence", 
                                       "Hypertension",
                                       "Diabetes"))
data_imputed$Province <- factor(data_imputed$Province, levels = c("15", "43", "44", "51", "52", "61", "62"), 
                                labels = c("Inner Mongolia Autonomous Region", 
                                           "Hunan Province",
                                           "Guangdong Province", 
                                           "Sichuan Province",
                                           "Guizhou Province", 
                                           "Shaanxi Province",
                                           "Gansu Province"))
data_imputed$Area <- factor(data_imputed$Area, levels = c("1", "2"), 
                                labels = c("Urban", 
                                           "Rural"))
data_imputed$Doctor.gender <- factor(data_imputed$Doctor.gender, levels = c("0", "1"), 
                            labels = c("Male", 
                                       "Female"))
data_imputed$Doctor.age <- factor(data_imputed$Doctor.age, levels = c("1", "2", "3"), 
                                     labels = c("<30", 
                                                "30-50", 
                                                "≥50"))
data_imputed$Doctor.ethnicity <- factor(data_imputed$Doctor.ethnicity, levels = c("1", "2", "9"), 
                                  labels = c("Han", 
                                             "Non-Han", 
                                             "Uncertain"))
data_imputed$Patient.gender <- factor(data_imputed$Patient.gender, levels = c("0", "1"), 
                                     labels = c("Male", 
                                                "Female"))



set.seed(25)
rf_model <- randomForest(
  as.factor(Ownership) ~ case + Province + Area + Doctor.gender + Doctor.age + Doctor.ethnicity + Patient.gender,
  data = data_imputed,
  ntree = 500,
  mtry = 2,
  importance = TRUE
)
data_imputed$propensity_score <- predict(rf_model, type = "prob")[, 2]

# 4 执行匹配
set.seed(2025)
matched_data <- matchit(
  Ownership ~ case + Province + Area + Doctor.gender + Doctor.age + + Doctor.ethnicity + Patient.gender,
  data = data_imputed,
  method = "nearest",
  distance = data_imputed$propensity_score,
  caliper = 0.25,
  replace = FALSE, #无放回
  ratio = 1                         # 1:1匹配增加对照样本
)
matched_df <- match.data(matched_data)

b1 <- bal.tab(Ownership ~ case + Province + Area + Doctor.gender + Doctor.age + 
                Doctor.ethnicity + Patient.gender,
              data = data_imputed,
              int = TRUE)
v1 <- var.names(b1, type = "vec", minimal = TRUE)
v1["case"] <- "Disease cases"
v1["Province"] <- "Province"
v1["Area"] <- "Area"
v1["Doctor.gender"] <- "Physician Gender"
v1["Doctor.age"] <- "Physician Age"
v1["Doctor.ethnicity"] <- "Physician Ethnicity"
v1["Patient.gender"] <- "Patient Gender"

love.plot(matched_data, threshold = 0.1,var.names = v1)

