# 加载必要的包
library(randomForest)
library(ggplot2)
library(dplyr)
library(caret)

# 准备数据
# 选择分析中使用的变量
predictors <- c("case", "Area", "Province", "Ownership", "Hospital.level", 
                "Institution.type", "Doctor.gender", "Doctor.age", 
                "Doctor.ethnicity", "Patient.gender")

# 确保所有预测变量都是因子类型
data_rf <- data0 %>%
  select(Treatment.score, all_of(predictors)) %>%
  mutate(across(all_of(predictors), as.factor)) %>%
  na.omit()  # 移除缺失值

# 设置随机种子以确保结果可重现
set.seed(2025)

# 训练随机森林模型
rf_model <- randomForest(
  Treatment.score ~ .,
  data = data_rf,
  ntree = 500,        # 树的数量
  mtry = 3,           # 每次分割时考虑的变量数
  importance = TRUE,  # 计算变量重要性
  proximity = FALSE
)

# 提取变量重要性
importance_df <- as.data.frame(importance(rf_model))
##不要importance_df$Variable <- rownames(importance_df)
importance_df$Variable <- c("Disease cases", 
                                   "Institution Type",                                    
                                   "Province", 
                                   "Physician Gender",
                                   "Patient Gender", 
                                   "Institution Level",
                                   "Area",
                                   "Physician Age",
                                   "Physician Ethnicity",
                                   "Ownership"
                                   )

rownames(importance_df) <- NULL

# 按重要性排序
importance_df <- importance_df %>%
  arrange(desc(`%IncMSE`))

# 创建变量重要性图
p <- ggplot(importance_df, aes(x = reorder(Variable, `%IncMSE`), y = `%IncMSE`)) +
  geom_bar(stat = "identity", fill = "steelblue") +
  coord_flip() +
  labs(
    title = "The characteristic importance graph (Treatment score)",
    subtitle = "Based on %IncMSE",
    x = "",
    y = "%IncMSE"
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
ggsave("特征重要性图.png", p, width = 10, height = 8, dpi = 300)

# 打印最重要的10个变量
cat("最重要的10个变量:\n")
print(head(importance_df, 10))

# 创建详细的结果表格
importance_table <- importance_df %>%
  select(Variable, `%IncMSE`, IncNodePurity) %>%
  rename(
    "变量" = Variable,
    "增加MSE百分比" = `%IncMSE`,
    "节点纯度增加" = IncNodePurity
  )

# 导出到Excel
wb <- createWorkbook()
addWorksheet(wb, "特征重要性")
writeData(wb, sheet = "特征重要性", importance_table)
saveWorkbook(wb, "随机森林特征重要性结果.xlsx", overwrite = TRUE)

# 显示图形
print(p)

# 输出模型性能
cat("\n模型性能:\n")
cat("均方误差 (MSE):", mean(rf_model$mse), "\n")
cat("方差解释率:", round(100 * (1 - mean(rf_model$mse) / var(data_rf$Treatment.score)), 2), "%\n")