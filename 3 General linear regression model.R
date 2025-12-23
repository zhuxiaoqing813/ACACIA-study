#一.多变量回归分析
##1.1 一般线性回归分析（治疗质量）
library(tidyverse)
library(openxlsx)

# 1. 数据准备
## 1.1 读取匹配前数据
data0 <- read.csv("D:/ACACIA 分析/分析数据/250708/Acacia原始数据250723-1800.csv")

### 定义变量类型（只保留需要的变量类型转换）
#### 因子变量
data0$case <- as.factor(data0$case)
data0$Area <- as.factor(data0$Area)
data0$Province <- as.factor(data0$Province)
data0$Ownership <- as.factor(data0$Ownership)
data0$Hospital.level <- as.factor(data0$Hospital.level)
data0$Institution.type <- as.factor(data0$Institution.type)
data0$Doctor.gender <- as.factor(data0$Doctor.gender)
data0$Doctor.age <- as.factor(data0$Doctor.age)
data0$Doctor.ethnicity <- as.factor(data0$Doctor.ethnicity)
data0$Patient.gender <- as.factor(data0$Patient.gender)

#### 数值变量（因变量）
data0$Treatment.score <- as.numeric(data0$Treatment.score)

# 设置参照组
data0$case <- relevel(data0$case, ref = "1")  # 将"偏头痛"设为参照组
data0$Area <- relevel(data0$Area, ref = "2")  # 将"乡村"设为参照组
data0$Province <- relevel(data0$Province, ref = "15")  # 将"内蒙古自治区"设为参照组
data0$Ownership <- relevel(data0$Ownership, ref = "0")  # 将"公立"设为参照组
data0$Hospital.level <- relevel(data0$Hospital.level, ref = "1")  # 将"一级"设为参照组
data0$Institution.type <- relevel(data0$Institution.type, ref = "1")  # 将"村卫生室"设为参照组
data0$Doctor.age <- relevel(data0$Doctor.age, ref = "1")  # 将"<30岁"设为参照组
data0$Doctor.ethnicity <- relevel(data0$Doctor.ethnicity, ref = "1")  # 将"汉族"设为参照组
data0$Patient.gender <- relevel(data0$Patient.gender, ref = "1")  # 将"女性"设为参照组

# 2. 建模（一般线性回归）
set.seed(2025)
lm_model <- lm(Treatment.score ~ 
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

# 3. 提取模型结果
# 获取模型摘要
model_summary <- summary(lm_model)

# 提取系数结果
coef_mat <- coef(model_summary)
results_table <- as.data.frame(coef_mat)

# 计算95%置信区间
ci_mat <- confint(lm_model)
results_table <- cbind(results_table, ci_mat)

# 重命名列
colnames(results_table) <- c("Estimate", "Std_Error", "t_value", "p_value", "CI_95_lower", "CI_95_upper")

# 添加显著性标记
results_table$Significance <- ifelse(results_table$p_value < 0.001, "***",
                                     ifelse(results_table$p_value < 0.01, "**",
                                            ifelse(results_table$p_value < 0.05, "*", "")))

# 将变量名作为单独列
results_table$Variable <- rownames(results_table)
rownames(results_table) <- NULL
results_table <- results_table %>% select(Variable, everything())

# 4. 创建模型摘要表
model_stats <- data.frame(
  Metric = c("R-squared", "Adjusted R-squared", "Residual Std. Error", "F-statistic", "p-value (F-test)", "DF Residual", "DF Model"),
  Value = c(
    model_summary$r.squared,
    model_summary$adj.r.squared,
    model_summary$sigma,
    model_summary$fstatistic[1],
    pf(model_summary$fstatistic[1], model_summary$fstatistic[2], model_summary$fstatistic[3], lower.tail = FALSE),
    model_summary$df[2],
    model_summary$df[1]
  )
)

# 5. 创建变量描述表
variable_description <- data.frame(
  Variable = c("Treatment.score", "case", "Area", "Province", "Ownership", 
               "Hospital.level", "Institution.type", "Doctor.gender", 
               "Doctor.age","Doctor.ethnicity", "Patient.gender"),
  Type = c("数值变量", "十一分类因子", "二分类因子", "七分类因子", "二分类因子",
           "三分类因子", "五分类因子", "二分类因子", "三分类因子", "三分类因子","二分类因子"),
  Reference = c("不适用", 
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
  Description = c("治疗质量评分（数值越高表示质量越好）", 
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

# 6. 导出到Excel
# 创建工作簿
wb <- createWorkbook()

# 添加工作表
addWorksheet(wb, "回归系数")
addWorksheet(wb, "模型摘要")
addWorksheet(wb, "变量描述")

# 写入数据
writeData(wb, sheet = "回归系数", results_table)
writeData(wb, sheet = "模型摘要", model_stats)
writeData(wb, sheet = "变量描述", variable_description)

# 设置数字格式
num_style <- createStyle(numFmt = "0.0000")
addStyle(wb, sheet = "回归系数", style = num_style, 
         rows = 2:(nrow(results_table) + 1), 
         cols = c(2, 3, 4, 5, 6, 7),
         gridExpand = TRUE)

# 保存Excel文件
output_file <- "一般线性回归分析_Treatment.score_new.xlsx"
saveWorkbook(wb, output_file, overwrite = TRUE)

# 打印成功消息
cat(paste0("分析结果已成功导出到: ", output_file, "\n"))
cat(paste0("工作目录: ", getwd(), "\n"))