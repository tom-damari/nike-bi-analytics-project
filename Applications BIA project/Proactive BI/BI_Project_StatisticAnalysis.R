library("readxl")
library(dplyr)
library(ggplot2)
library(car)
library(multcompView)
library(corrplot)

install.packages("corrplot")

file_path <- "C:/Users/User/Desktop/תואר שני/שנה ד סמסטר ב/בינה ואנטיקה עסקית/פרוייקט BI/Nike Database_Final.xlsx"
df <- read_excel(file_path, sheet = "FACT - INCLUDES")

############################## ANOVA - Income Per Season ##############################
df$`Order Date` <- as.Date(df$`Order Date`)

df$Season <- ifelse(format(as.Date(df$`Order Date`), "%m") %in% c("12", "01", "02"), "Winter",
                    ifelse(format(as.Date(df$`Order Date`), "%m") %in% c("03", "04", "05"), "Spring",
                           ifelse(format(as.Date(df$`Order Date`), "%m") %in% c("06", "07", "08"), "Summer", "Fall")))
seasonal_sales <- df %>%
  group_by(Season) %>%
  summarize(MeanSales = mean(Income))

boxplot(Income ~ Season, data = df, main = "Sales Distribution by Season",
        xlab = "Season", ylab = "Sales Income", col = "lightblue")

anova_model <- aov(Income ~ Season, data = df)
summary(anova_model)


############################## Corolation - Quantity By Color ##############################

summary_df <- df %>%
  group_by(`Product Name`) %>%
  summarize(AveragePrice = mean(`Version Price`, na.rm = TRUE),
            TotalQuantity = sum(Quantity, na.rm = TRUE))

print(head(summary_df))

correlation_result <- cor(summary_df$AveragePrice, summary_df$TotalQuantity, use = "complete.obs")
print(paste("Correlation between Average Price and Total Quantity:", correlation_result))

cor_matrix <- cor(summary_df[, c("AveragePrice", "TotalQuantity")], use = "complete.obs")
print(cor_matrix)

corrplot(cor_matrix, method = "color", col = colorRampPalette(c("blue", "white", "red"))(200),
         addCoef.col = "black", 
         tl.col = "black", tl.srt = 45, 
         number.cex = 0.8, 
         cl.cex = 0.8) 

ggplot(summary_df, aes(x = AveragePrice, y = TotalQuantity)) +
  geom_point(color = "blue") +
  geom_smooth(method = "lm", se = FALSE, col = "red") +
  labs(title = "Orders Amount vs. Product Average Price", x = "Product Average Price", y = "Orders Amount") +
  theme(axis.text.x = element_text(angle = 45, hjust = 1))




#################################


customers <- read_excel(file_path, sheet = "CUSTOMERS")
sales_data <- merge(df, customers, by.x = "Customer Email", by.y = "Customer Email")

contingency_table <- table(sales_data$Color, sales_data$`Customer Gender`)

income_by_age <- sales_data %>%
  group_by(`Cusomer Age`, `Customer Gender`) %>%
  summarize(TotalIncome = sum(Income, na.rm = TRUE))

t_test_result <- t.test(Income ~ `Customer Gender`, data = sales_data)
print(t_test_result)

correlation_result <- cor(income_by_age$`Cusomer Age`, income_by_age$TotalIncome, use = "complete.obs")
print(paste("Correlation between Customer Age and Total Income:", correlation_result))

ggplot(income_by_age, aes(x = `Cusomer Age`, y = TotalIncome, color = `Customer Gender`)) +
  geom_point() +
  geom_smooth(method = "lm", se = FALSE, col = "red") +
  labs(title = "Total Income vs. Customer Age", x = "Customer Age", y = "Total Income") +
  theme_minimal()
