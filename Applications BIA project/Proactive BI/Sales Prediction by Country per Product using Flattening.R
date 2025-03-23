# Load libraries
library(openxlsx)
library(dplyr)
library(lubridate)
library(caret)
library(ggplot2)
library(car)
library(lmtest)
library(MASS)
library(mgcv)

# Load data
file_path <- file.choose()
CUSTOMERS <- read.xlsx(file_path, sheet = "CUSTOMERS")
INCLUDES <- read.xlsx(file_path, sheet = "FACT - INCLUDES")

# Merge INCLUDES with CUSTOMERS to get country information
INCLUDES_CUSTOMERS <- INCLUDES %>%
  inner_join(CUSTOMERS, by = c("Customer.Email" = "Customer.Email"))

# Convert Order.Date to date format with origin
INCLUDES_CUSTOMERS <- INCLUDES_CUSTOMERS %>%
  mutate(Order.Date = as.Date(Order.Date, origin = "2006-04-01")) 

# Extract year from the order date
INCLUDES_CUSTOMERS <- INCLUDES_CUSTOMERS %>%
  mutate(Order.Year = year(Order.Date))

# Convert categorical variables to factors before flattening
INCLUDES_CUSTOMERS <- INCLUDES_CUSTOMERS %>%
  mutate(Customer_Gender = as.factor(Customer.Gender),
         Customer_Age_Group = as.factor(Customer.Age.Group),
         Has_Membership = as.factor(Has.Membership),
         Order_Type = as.factor(Order.Type))


# Function to find the most frequent value
most_frequent <- function(x) {
  ux <- unique(x)
  ux[which.max(tabulate(match(x, ux)))]
}

# Group by Address-Country, Product.Name, and Order.Year, then summarize data
annual_data <- INCLUDES_CUSTOMERS %>%
  group_by(`Address-Country`, `Product.Name`, Order.Year) %>%
  summarise(Annual_Quantity = sum(Quantity),
            Avg_Price = mean(Version.Price, na.rm = TRUE),
            Avg_Age = mean(Cusomer.Age, na.rm = TRUE),
            Customer_Age_Group = most_frequent(Customer_Age_Group),
            Most_Frequent_Gender = most_frequent(Customer_Gender),
            Most_Frequent_Order_Type = most_frequent(Order_Type),
            Most_Frequent_Membership = most_frequent(Has.Membership),
            Male_Count = sum(Customer_Gender == "Male"),
            Online_Count = sum(Order_Type == "Online"),
            Membership_Count = sum(Has.Membership == "Y"),
            .groups = 'drop')


# Function to calculate the slope
calc_slope <- function(x) {
  time <- 1:length(x)
  lm(x ~ time)$coefficients[2]
}

# Aggregate to get the slope, min, max, and average for each year
data <- annual_data %>%
  group_by(`Address-Country`, `Product.Name`) %>%
  summarise(Avg_Annual_Quantity = mean(Annual_Quantity, na.rm = TRUE),
            Slope_Annual_Quantity = calc_slope(Annual_Quantity),
            Min_Annual_Quantity = min(Annual_Quantity),
            Max_Annual_Quantity = max(Annual_Quantity),
            Avg_Price = mean(Avg_Price, na.rm = TRUE),
            Avg_Age = mean(Avg_Age, na.rm = TRUE),
            Customer_Age_Group = most_frequent(Customer_Age_Group),
            Most_Frequent_Gender = most_frequent(Most_Frequent_Gender),
            Most_Frequent_Order_Type = most_frequent(Most_Frequent_Order_Type),
            Most_Frequent_Membership = most_frequent(Most_Frequent_Membership),
            Avg_Male_Count = mean(Male_Count, na.rm = TRUE),
            Avg_Online_Count = mean(Online_Count, na.rm = TRUE),
            Avg_Membership_Count = mean(Membership_Count, na.rm = TRUE),
            .groups = 'drop')

# Handle missing values
handle_missing_values <- function(data) {
  data %>%
    mutate(across(where(is.numeric), ~ ifelse(is.na(.), 0, .))) %>%
    mutate(across(where(is.factor), ~ ifelse(is.na(.), most_frequent(.), .)))
}

# Apply the missing values handling function
data <- handle_missing_values(data)

# Remove columns with only NA values
data <- data[, colSums(!is.na(data)) > 0]

############################################ LINEAR REGRESSION MODEL ############################################

# Split by 80-20 Train and Test sets
set.seed(123)
trainIndex <- createDataPartition(data$Avg_Annual_Quantity, p = 0.8, list = FALSE)
trainData <- data[trainIndex, ]
testData <- data[-trainIndex, ]

# Train a regression model excluding 'Address-Country' and 'Product.Name'
model <- train(Avg_Annual_Quantity ~ Slope_Annual_Quantity + Min_Annual_Quantity + Max_Annual_Quantity + 
                 Avg_Price + Avg_Age + factor(Customer_Age_Group) + factor(Most_Frequent_Gender) + 
                 factor(Most_Frequent_Order_Type) + factor(Most_Frequent_Membership) + 
                 Avg_Male_Count + Avg_Online_Count + Avg_Membership_Count, 
               data = trainData, method = "lm")

# Get the summary of the model
model_summary <- summary(model$finalModel)
print(model_summary)

# Calculate AIC and BIC
aic_value <- AIC(model$finalModel)
bic_value <- BIC(model$finalModel)
print(paste("AIC:", aic_value, "     BIC:", bic_value))

# Make predictions
predictions <- predict(model, testData)

# Round predictions up to the nearest integer
rounded_predictions <- ceiling(predictions)

# Output predictions
results <- data.frame(
  Country = testData$`Address-Country`,   
  Product_Name = testData$`Product.Name`,  
  Actual_Quantity = testData$Avg_Annual_Quantity,
  Predicted_Quantity = rounded_predictions
)

# Print Results
print(results)

# Calculate RMSE
rmse <- sqrt(mean((results$Actual_Quantity - results$Predicted_Quantity)^2))
cat("Root Mean Squared Error (RMSE):", rmse, "\n")

############################################ MULTICOLINERITY ############################################

# Variance Inflation Factor (VIF) for Multicolinearity
vif_values <- vif(model$finalModel)
vif_table <- data.frame(VIF_Value = vif_values)
print(vif_table)

############################################ LINEAR REGRESSION ASSUMPTIONS ############################################

# Residual Analysis
residuals <- residuals(model$finalModel)

# Residuals vs. Fitted Values Plot
ggplot(data = data.frame(Fitted = fitted(model$finalModel), Residuals = residuals), aes(x = Fitted, y = Residuals)) +
  geom_point() +
  geom_hline(yintercept = 0, col = "red") +
  labs(title = "Residuals vs Fitted Values", x = "Fitted Values", y = "Residuals")

# Breusch-Pagan Test for Homoscedasticity Assumption
bptest(model$finalModel)

# Durbin-Watson Test for Independence Assumption
dwtest(model$finalModel)

# Histogram of Residuals
ggplot(data = data.frame(Residuals = residuals), aes(x = Residuals)) +
  geom_histogram(binwidth = 1, fill = "blue", color = "black", alpha = 0.7) +
  labs(title = "Histogram of Residuals", x = "Residuals", y = "Frequency")

# Shapiro-Wilk Test for Normality
shapiro_test <- shapiro.test(residuals)
print(shapiro_test)

# KS Test for Normality
ks_test <- ks.test(residuals, "pnorm", mean = mean(residuals), sd = sd(residuals))
print(ks_test)

# Q-Q Plot of Residuals with Test Results
qq_plot <- ggplot(data.frame(sample = residuals), aes(sample = sample)) +
  stat_qq() +
  stat_qq_line() +
  labs(title = "Q-Q Plot of Residuals")
print(qq_plot)


