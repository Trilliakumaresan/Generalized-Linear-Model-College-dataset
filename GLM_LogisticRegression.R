#Name- Aquila Persis Pillay, Date- 29/01/2024
cat("\014") # clears console
rm(list = ls()) # clears global environment
try(dev.off(dev.list()["RStudioGD"]), silent = TRUE) # clears plots
# Load required libraries
library(readxl)
library(ggplot2)
library(tidyverse)
library(forecast)

# Load the data
excel_file_path <- "Project_Data.xlsx"
stock_prices <- readxl::read_excel(excel_file_path)
# Renaming columns
names(stock_prices) <- c("Date", "Period", "AAPL_Price", "AAPL_Volume", "HON_Price", "HON_Volume")
calculate_wma <- function(series, weights) {
  n <- length(series)
  
  if (length(weights) != 3) {
    stop("Weights vector must have three elements.")
  }
  
  wma <- rep(NA, n)
  
  for (i in 4:n) {
    wma[i] <- sum(weights * series[(i-2):(i-1)])
  }
  
  return(wma)
}

# Function to calculate weighted moving average percentage error (WMAPE)
calculate_wmape <- function(actuals, forecast, weights) {
  n <- length(actuals)
  wmape <- sum(weights * abs(actuals - forecast) / abs(actuals)) / sum(weights) * 100
  return(wmape)
}


# Converting the 'Date' column to a Date format
stock_prices$Date <- as.Date(stock_prices$Date, format="%m/%d/%Y")
# Part 1: Short-term Forecasting

# Question 1
# Create a simple line plot
ggplot(stock_prices, aes(x = Date)) +
  geom_line(aes(y = AAPL_Price, color = "AAPL"), size = 1) +
  geom_line(aes(y = HON_Price, color = "HON"), size = 1) +
  labs(title = "Stock Prices Over Time",
       x = "Date",
       y = "Stock Price",
       color = "Stock") +
  theme_minimal()

# Question 2
# AAPL exponential smoothing
aapl_ts <- ts(stock_prices$AAPL_Price , frequency = 252)
aapl_ts <- na.omit(aapl_ts)
aapl_ts <- aapl_ts[is.finite(aapl_ts)]


train_data <- window(aapl_ts, end = c(252, 1))
test_data <- window(aapl_ts, start = c(1))
calculate_mape <- function(alpha_value) {
  model <- HoltWinters(train_data, alpha = alpha_value, beta = FALSE, gamma = FALSE)
  forecast_values <- forecast(model, h = 1)$mean
  mape <- mean(abs((test_data - forecast_values) / test_data) * 100)
  cat("Alpha:", alpha_value, "MAPE:", mape, "\n")
  return(list(model = model, forecast_values = forecast_values, mape = mape))
}

alpha_values <- c(0.15, 0.35, 0.55, 0.75)
mape_results <- sapply(alpha_values, calculate_mape)

print(mape_results)
best_alpha_index <- which.min(mape_results[3, ])
best_alpha_aapl <- alpha_values[best_alpha_index]
mape_best_alpha_aapl <- mape_results[3, best_alpha_index]
best_model_aapl <- mape_results$model[[best_alpha_index]]
cat("Best Alpha for AAPL:", best_alpha_aapl, "\n")


# Question 3
# HON exponential smoothing
hon_ts <- ts(stock_prices$HON_Price, frequency = 252)
hon_ts <- na.omit(hon_ts)
hon_ts <- hon_ts[is.finite(hon_ts)]


train_data_hon <- window(hon_ts, end = c(252, 1))
test_data_hon <- window(hon_ts, start = c(1))

calculate_mape_hon <- function(alpha_value) {
  model_hon <- HoltWinters(train_data_hon, alpha = alpha_value, beta = FALSE, gamma = FALSE)
  forecast_values_hon <- forecast(model_hon, h = 1)$mean
  mape_hon <- mean(abs((test_data_hon - forecast_values_hon) / test_data_hon) * 100)
  return(mape_hon)
}

alpha_values_hon <- c(0.15, 0.35, 0.55, 0.75)
mape_values_hon <- sapply(alpha_values_hon, calculate_mape_hon)
best_alpha_hon <- alpha_values_hon[which.min(mape_values_hon)]
mape_best_alpha_hon <- calculate_mape_hon(best_alpha_hon)

cat("HON MAPE Values:", mape_values_hon, "\n")
cat("Best Alpha for HON:", best_alpha_hon, "\n")
cat("MAPE with Best Alpha for HON:", mape_best_alpha_hon, "\n")

# Question 4
# AAPL exponential smoothing with trend
calculate_mape_trend_aapl <- function(beta_value) {
  model_trend_aapl <- HoltWinters(train_data, alpha = 0.55, beta = beta_value, gamma = FALSE)
  forecast_values_trend_aapl <- forecast(model_trend_aapl, h = 1)$mean
  mape_trend_aapl <- mean(abs((test_data - forecast_values_trend_aapl) / test_data) * 100)
  return(mape_trend_aapl)
}

beta_values_aapl <- c(0.15, 0.25, 0.45, 0.85)
mape_values_trend_aapl <- sapply(beta_values_aapl, calculate_mape_trend_aapl)
best_beta_aapl <- beta_values_aapl[which.min(mape_values_trend_aapl)]
mape_best_beta_aapl <- calculate_mape_trend_aapl(best_beta_aapl)

cat("AAPL Trend MAPE Values:", mape_values_trend_aapl, "\n")
cat("Best Beta for AAPL Trend:", best_beta_aapl, "\n")
cat("MAPE with Best Beta for AAPL Trend:", mape_best_beta_aapl, "\n")

# Question 5
# HON exponential smoothing with trend
calculate_mape_trend_hon <- function(beta_value) {
  model_trend_hon <- HoltWinters(train_data_hon, alpha = 0.55, beta = beta_value, gamma = FALSE)
  forecast_values_trend_hon <- forecast(model_trend_hon, h = 1)$mean
  mape_trend_hon <- mean(abs((test_data_hon - forecast_values_trend_hon) / test_data_hon) * 100)
  return(mape_trend_hon)
}

beta_values_hon <- c(0.15, 0.25, 0.45, 0.85)
mape_values_trend_hon <- sapply(beta_values_hon, calculate_mape_trend_hon)
best_beta_hon <- beta_values_hon[which.min(mape_values_trend_hon)]
mape_best_beta_hon <- calculate_mape_trend_hon(best_beta_hon)

cat("HON Trend MAPE Values:", mape_values_trend_hon, "\n")
cat("Best Beta for HON Trend:", best_beta_hon, "\n")
cat("MAPE with Best Beta for HON Trend:", mape_best_beta_hon, "\n")

# Part 2: Long-term Forecasting

# Question 6
# AAPL 3-period weighted moving averages and linear trend

# Given weights
weights <- c(0.5, 0.3, 0.2)

# AAPL stock prices
aapl_series <- stock_prices$AAPL_Price


# Calculate WMA for AAPL
wma_aapl <- calculate_wma(aapl_series[1:100], weights)

# Linear trend for periods 1 to 252
time_periods_aapl <- 1:length(aapl_series[1:252])
linear_trend_aapl <- lm(aapl_series[1:252] ~ time_periods_aapl)

# Extract coefficients for AAPL
intercept_aapl <- coef(linear_trend_aapl)[1]
slope_aapl <- coef(linear_trend_aapl)[2]

# Generate linear trend values for periods 101 to 252
linear_trend_values_aapl <- intercept_aapl + slope_aapl * (time_periods_aapl + 100)

# Combine WMA and Linear Trend for periods 101 to 252
combined_forecast_aapl <- c(wma_aapl, linear_trend_values_aapl)

# Actual values for periods 1 to 252
actual_values_aapl <- aapl_series[1:252]

# Calculate MAPE for periods 1 to 252
mape_aapl <- calculate_wmape(actual_values_aapl, combined_forecast_aapl, weights)

cat("MAPE for periods 1 to 252:", round(mape_aapl, 2), "%\n")


# Question 7
# AAPL linear regression forecast for periods 253 to 257
# Given data
aapl_series <- stock_prices$AAPL_Price
time_periods <- 1:length(aapl_series)

# Performing linear regression
linear_model_aapl <- lm(cbind(aapl_series) ~ time_periods)

# Generating future time periods for forecasting
future_time_periods_aapl <- 253:257

# Predicting using the linear model
linear_reg_forecast_values_aapl <- predict(linear_model_aapl, newdata = data.frame(time_periods = future_time_periods_aapl))

# Displaying the forecast values
cat("Linear Regression Forecast for AAPL (Periods 253 to 257):", linear_reg_forecast_values_aapl, "\n")


#Question 8
# HON 3-period weighted moving averages and linear trend
weights <- c(0.5, 0.3, 0.2)
hon_series <- stock_prices$HON_Price 

# Calculate WMA for HON
wma_hon <- calculate_wma(hon_series, weights)

# Corrected Linear Trend Calculation for HON
time_periods_hon <- 1:length(hon_series)

# Linear trend for HON
linear_trend_hon <- lm(cbind(hon_series) ~ time_periods_hon)

# Extract coefficients for HON
intercept_hon <- coef(linear_trend_hon)[1]
slope_hon <- coef(linear_trend_hon)[2]

# Generate linear trend values for HON
linear_trend_values_hon <- intercept_hon + slope_hon * time_periods_hon

# Combine WMA and Linear Trend for HON
combined_forecast_hon <- wma_hon
combined_forecast_hon[101:252] <- linear_trend_values_hon[101:252]

# Calculate WMAPE for the combined forecast for HON
mape_wma_hon <- calculate_wmape(hon_series[101:252], combined_forecast_hon[101:252], weights)

cat("WMAPE for HON:", mape_wma_hon, "\n")

# Question 9
# HON linear regression forecast for periods 253 to 257

# Extracting the relevant data
hon_series <- stock_prices$HON_Price 

# Creating a time variable
time_periods_hon <- 1:length(hon_series)

# Performing linear regression
linear_model_hon <- lm(cbind(hon_series) ~ time_periods_hon)

# Generating future time periods for forecasting
future_time_periods_hon <- 253:257

# Predicting using the linear model
linear_reg_forecast_values_hon <- predict(linear_model_hon, newdata = data.frame(time_periods_hon = future_time_periods_hon))

# Displaying the forecast values
cat("Linear Regression Forecast for HON (Periods 253 to 257):", linear_reg_forecast_values_hon, "\n")

# Part 3: Regression

# Question 10
# AAPL simple regression

time_periods_aapl <- 1:length(aapl_series)

linear_model_aapl <- lm(cbind(aapl_series) ~ time_periods_aapl)

future_time_periods_aapl <- 1:257

linear_reg_forecast_values_aapl <- predict(linear_model_aapl, newdata = data.frame(time_periods_aapl = future_time_periods_aapl))
forecast_values_1_to_252_aapl <- linear_reg_forecast_values_aapl[1:252]

# Question 11: HON simple regression
hon_series <- stock_prices$HON_Price

time_periods_hon <- 1:length(hon_series)

linear_model_hon <- lm(cbind(hon_series) ~ time_periods_hon)

future_time_periods_hon <- 1:257

linear_reg_forecast_values_hon <- predict(linear_model_hon, newdata = data.frame(time_periods_hon = future_time_periods_hon))

forecast_values_1_to_252_hon <- linear_reg_forecast_values_hon[1:252]
