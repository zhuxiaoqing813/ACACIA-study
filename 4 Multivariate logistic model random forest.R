# 加载必要的包
library(randomForest)
library(ggplot2)
library(dplyr)
library(caret)

# 准备数据 - 使用与多元Logistic回归相同的变量
predictors <- c("case", "Area", "Province", "Ownership", "Hospital.level", 
                "Institution.type", "Doctor.gender", "Doctor.age", 
                "Doctor.ethnicity", "Patient.gender")

# 确保因变量是因子类型
data_rf_diagnosis <- data0 %>%
  select(Diagnosis.quality, all_of(predictors)) %>%
  mutate(across(all_of(c("Diagnosis.quality", predictors)), as.factor)) %>%
  na.omit()  # 移除缺失值

# 检查数据平衡性
table(data_rf_diagnosis$Diagnosis.quality)

# 设置随机种子以确保结果可重现
set.seed(2025)

# 训练随机森林模型（分类问题）
rf_model_diagnosis <- randomForest(
  Diagnosis.quality ~ .,
  data = data_rf_diagnosis,
  ntree = 500,        # 树的数量
  mtry = floor(sqrt(length(predictors))), # 分类问题常用的mtry设置
  importance = TRUE,  # 计算变量重要性
  proximity = FALSE
)

# 提取变量重要性
importance_diagnosis <- as.data.frame(importance(rf_model_diagnosis))
##不要importance_diagnosis$Variable <- rownames(importance_diagnosis)
importance_diagnosis$Variable <- c("Disease cases", 
                                   "Province", 
                                   "Patient Gender", 
                                   "Institution Type", 
                                   "Physician Ethnicity",
                                   "Institution Level",
                                   "Physician Gender",
                                   "Area",
                                   "Ownership",
                                   "Physician Age")



# 按MeanDecreaseAccuracy排序（分类问题的主要重要性指标）
importance_diagnosis <- importance_diagnosis %>%
  arrange(desc(MeanDecreaseAccuracy))

# 创建变量重要性图
p_diagnosis <- ggplot(importance_diagnosis, aes(x = reorder(Variable, MeanDecreaseAccuracy), 
                                                y = MeanDecreaseAccuracy)) +
  geom_bar(stat = "identity", fill = "coral", alpha = 0.8) +
  coord_flip() +
  labs(
    title = "The characteristic importance graph (Diagnostic quality)",
    subtitle = "Based on Mean Decrease Accuracy",
    x = "",
    y = "Mean Decrease Accuracy"
  ) +
  theme_minimal(base_size = 14) +
  theme(
    plot.title = element_text(hjust = 1, face = "bold", size = 14),
    plot.subtitle = element_text(hjust = 0.5, size = 12),
    axis.text.y = element_text(size = 12, face = "bold"),
    axis.text.x = element_text(size = 12),
    axis.title.x = element_text(size = 14)
  ) +
  scale_y_continuous(expand = expansion(mult = c(0, 0.05)))

# 保存图形
ggsave("诊断质量特征重要性图.png", p_diagnosis, width = 12, height = 9, dpi = 300, bg = "white")

# 显示图形
print(p_diagnosis)

# 创建详细的结果表格
importance_table_diagnosis <- importance_diagnosis %>%
  select(Variable, MeanDecreaseAccuracy, MeanDecreaseGini) %>%
  rename(
    "变量" = Variable,
    "平均精度下降" = MeanDecreaseAccuracy,
    "基尼指数下降" = MeanDecreaseGini
  )

# 导出到Excel
wb <- createWorkbook()
addWorksheet(wb, "诊断质量特征重要性")
writeData(wb, sheet = "诊断质量特征重要性", importance_table_diagnosis)
saveWorkbook(wb, "诊断质量随机森林特征重要性结果.xlsx", overwrite = TRUE)

# 输出最重要的变量
cat("诊断质量最重要的10个变量:\n")
print(head(importance_diagnosis, 10))

# 输出模型性能
cat("\n模型性能:\n")
print(rf_model_diagnosis$confusion)
cat("整体准确率:", sum(diag(rf_model_diagnosis$confusion))/sum(rf_model_diagnosis$confusion), "\n")