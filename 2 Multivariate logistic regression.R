# 一.多变量回归分析
##1.1 多元Logistic回归分析（诊断质量）
library(tidyverse)
library(nnet)
library(caret)
library(openxlsx)
library(pROC)  #用于计算AUC

# 1. 数据准备
## 1.1 读取匹配前数据
data0 <- read.csv("D:/ACACIA 分析/分析数据/250708/Acacia原始数据250723-1800.csv")

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

# 设置参照组！
data0$Diagnosis.quality <- relevel(data0$Diagnosis.quality, ref = "4")  # 将"错误"设为参照组
data0$case <- relevel(data0$case, ref = "1")  # 将"偏头痛"设为参照组
data0$Area <- relevel(data0$Area, ref = "2")  # 将"乡村"设为参照组
data0$Province <- relevel(data0$Province, ref = "15")  # 将"内蒙古自治区"设为参照组
data0$Ownership <- relevel(data0$Ownership, ref = "0")  # 将"公立"设为参照组
data0$Hospital.level <- relevel(data0$Hospital.level, ref = "1")  # 将"一级"设为参照组
data0$Institution.type <- relevel(data0$Institution.type, ref = "1")  # 将"村卫生室"设为参照组
data0$Doctor.age <- relevel(data0$Doctor.age, ref = "1")  # 将"<30岁"设为参照组
data0$Doctor.ethnicity <- relevel(data0$Doctor.ethnicity, ref = "1")  # 将"汉族"设为参照组
data0$Patient.gender <- relevel(data0$Patient.gender, ref = "1")  # 将"女性"设为参照组

# 2. 建模
set.seed(2025)
multi_model <- multinom(Diagnosis.quality ~ 
                          case + 
                          Area + 
                          Province + 
                          Ownership + 
                          Hospital.level + 
                          Institution.type + 
                          Doctor.gender + 
                          Doctor.age + 
                          Doctor.ethnicity +
                          Patient.gender,
                        data = data0)

# 【核心新增部分开始：计算模型整体指标】
# 3.1 计算空模型（用于似然比检验和伪R²）
null_model <- multinom(Diagnosis.quality ~ 1, data = data0)

# 3.2 手动计算似然比检验
logLik_null <- logLik(null_model)
logLik_full <- logLik(multi_model)

# 计算似然比统计量
lr_stat <- 2 * (as.numeric(logLik_full) - as.numeric(logLik_null))

# 计算自由度
df_full <- attr(logLik_full, "df")
df_null <- attr(logLik_null, "df")
df_diff <- df_full - df_null

# 计算p值
p_value_lrt <- pchisq(lr_stat, df_diff, lower.tail = FALSE)

# 3.3 计算伪R²（Nagelkerke's）
nagelkerke_R2 <- (1 - exp((2/nrow(data0)) * (logLik_null - logLik_full))) / 
  (1 - exp(2 * logLik_null / nrow(data0)))

# 3.4 计算多分类AUC（宏观平均）
pred_probs <- fitted(multi_model)  # 使用 fitted() 而非 predict()

# 获取模型实际使用的响应变量（与 pred_probs 行数一致）
actual_response <- model.response(model.frame(multi_model))

auc_values <- numeric(ncol(pred_probs))
for(i in 1:ncol(pred_probs)){
  # 创建二分响应变量：当前类别 vs 其他
  binary_response <- as.numeric(actual_response == colnames(pred_probs)[i])
  roc_obj <- roc(response = binary_response,
                 predictor = pred_probs[,i])
  auc_values[i] <- auc(roc_obj)
}
mean_auc <- mean(auc_values)

# 3.5 打印核心指标到控制台
cat("\n======= 核心模型指标 =======\n")
cat("样本量 N =", nrow(data0), "\n")
cat("似然比检验: χ²(", df_diff, ") = ", 
    round(lr_stat, 2), ", p = ", 
    format.pval(p_value_lrt, digits = 4, eps = 0.0001), "\n", sep = "")
cat("伪R² (Nagelkerke's) =", round(nagelkerke_R2, 4), "\n")
cat("平均AUC (宏观) =", round(mean_auc, 4), "\n")
cat("AIC =", round(AIC(multi_model), 2), "\n")
cat("=============================\n")
# 【核心新增部分结束】

# 4. 提取模型结果
# 提取系数矩阵
coef_mat <- coef(summary(multi_model))
# 提取标准误矩阵
se_mat <- summary(multi_model)$standard.errors

# 计算z值和p值
z_values <- coef_mat / se_mat
p_values <- 2 * pnorm(-abs(z_values))  # 双侧检验

# 计算OR和95%CI
OR_values <- exp(coef_mat)
CI_lower <- exp(coef_mat - 1.96 * se_mat)
CI_upper <- exp(coef_mat + 1.96 * se_mat)

# 创建结果数据框
results_table <- data.frame(
  Predictor = rep(rownames(coef_mat), each = ncol(coef_mat)),
  Comparison = rep(colnames(coef_mat), times = nrow(coef_mat)),
  Beta = as.vector(t(coef_mat)),  # 转置后向量化
  SE = as.vector(t(se_mat)),
  OR = as.vector(t(OR_values)),
  OR_95CI_Lower = as.vector(t(CI_lower)),
  OR_95CI_Upper = as.vector(t(CI_upper)),
  z_value = as.vector(t(z_values)),
  p_value = as.vector(t(p_values))
)

# 添加显著性标记
results_table$Significance <- ifelse(results_table$p_value < 0.001, "***",
                                     ifelse(results_table$p_value < 0.01, "**",
                                            ifelse(results_table$p_value < 0.05, "*", "")))

# 5. 创建模型摘要表（包含新增指标）
model_summary <- data.frame(
  Metric = c("样本量 (N)", 
             "似然比检验 χ²(df)", 
             "似然比检验 p值", 
             "伪R² (Nagelkerke's)", 
             "平均AUC (宏观)", 
             "AIC",
             "残差偏差",
             "观测值数量"),
  Value = c(
    nrow(data0),
    paste0(round(lr_stat, 2), "(", df_diff, ")"),
    format.pval(p_value_lrt, digits = 4, eps = 0.0001),
    round(nagelkerke_R2, 4),
    round(mean_auc, 4),
    round(AIC(multi_model), 2),
    round(deviance(multi_model), 2),
    nrow(data0)
  )
)

# 6. 创建变量描述表
variable_description <- data.frame(
  Variable = c("Diagnosis.quality", "case", "Area", "Province", "Ownership", 
               "Hospital.level", "Institution.type", "Doctor.gender", 
               "Doctor.age","Doctor.ethnicity", "Patient.gender"),
  Type = c("五分类因子", "十一分类因子", "二分类因子", "七分类因子", "二分类因子",
           "三分类因子", "五分类因子", "二分类因子", "三分类因子", "三分类因子","二分类因子"),
  Reference = c("4=错误", 
                "1=偏头痛", 
                "2=乡村", 
                "15=内蒙古自治区", 
                "0=公立", 
                "1=一级", 
                "1=村卫生室", 
                "0=男", 
                "1=<30岁", 
                "1=汉族", 
                "1=女"),
  Levels = c("0=没有下诊断,1=完全正确,2=基本正确,3=部分正确,4=错误", 
             "1=偏头痛,2=产后抑郁,3=小儿腹泻,4=普通感冒,5=肺结核,6=胃炎,7=哮喘,8=心绞痛,9=腰背痛,10=压力性尿失禁,11=高血压,12=糖尿病", 
             "1=城镇,2=乡村",
             "15=内蒙古自治区,43=湖南省,44=广东省,51=四川省,52=贵州省,61=陕西省,62=甘肃省",
             "0=公立,1=民营",
             "1=一级,2=二级,9=未定级",
             "1=村卫生室,2=诊所,3=乡镇卫生院,4=社区卫生服务中心（站）,5=医院",
             "0=男,1=女",
             "1=<30岁,2=30-50岁,3=≥50岁",
             "1=汉族,2=非汉族,9=无法判断",
             "0=男,1=女")
)

# 7. 导出到Excel
# 创建工作簿
wb <- createWorkbook()

# 添加工作表
addWorksheet(wb, "回归结果")
addWorksheet(wb, "模型摘要")
addWorksheet(wb, "变量描述")

# 写入数据
writeData(wb, sheet = "回归结果", results_table)
writeData(wb, sheet = "模型摘要", model_summary)
writeData(wb, sheet = "变量描述", variable_description)

# 设置格式
# 设置OR值的数字格式
or_style <- createStyle(numFmt = "0.000")
addStyle(wb, sheet = "回归结果", style = or_style, 
         rows = 2:(nrow(results_table) + 1), 
         cols = c(4, 5, 6, 7),
         gridExpand = TRUE)

# 设置p值列的数字格式
p_value_style <- createStyle(numFmt = "0.0000")
addStyle(wb, sheet = "回归结果", style = p_value_style, 
         rows = 2:(nrow(results_table) + 1), 
         cols = 9,
         gridExpand = TRUE)

# 保存Excel文件
saveWorkbook(wb, "多元Logistic回归分析结果251223_new.xlsx", overwrite = TRUE)

# 添加模型系数表（可选）
coef_df <- as.data.frame(coef(multi_model))
coef_df$Predictor <- rownames(coef_df)
coef_df <- coef_df %>% select(Predictor, everything())

se_mat_df <- as.data.frame(summary(multi_model)$standard.errors)
se_mat_df$Predictor <- rownames(se_mat_df)
se_mat_df <- se_mat_df %>% select(Predictor, everything())

# 添加到新工作表
addWorksheet(wb, "模型beta系数")
writeData(wb, sheet = "模型beta系数", coef_df)
addWorksheet(wb, "模型SE系数")
writeData(wb, sheet = "模型SE系数", se_mat_df)
saveWorkbook(wb, "多元Logistic回归分析结果251223_new.xlsx", overwrite = TRUE)

# 打印成功消息
cat("分析结果已成功导出到: 多元Logistic回归分析结果251223_new.xlsx\n")
cat("工作目录:", getwd(), "\n")