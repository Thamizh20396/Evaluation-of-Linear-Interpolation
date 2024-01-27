library(openxlsx)
library(dplyr)
library(zoo)
library(readxl)
# Read Excel file into a dataframe
# Replace "E:/IODP 362/linear interpolation R.xlsx" with the actual file path
# The imported Excel file comprises two columns, with the first column labeled 'Age(Ma)' and the second column labeled 'MAR TOC (gC/cm2/k.y.)'.
data_new1 <- read_excel("E:/IODP 362/linear interpolation R.xlsx")
# Create 100 new columns (d_1, d_2, ..., d_100) in data_new1
# Each new column contains the same values from the column 'MAR TOC (gC/cm2/k.y.)'
for (i in 1:100) {
  # Generate column name
  col_name <- paste0("d_", i)
  # Assign values from 'MAR TOC (gC/cm2/k.y.)' to the new column
  data_new1[[col_name]] <- data_new1$`MAR TOC (gC/cm2/k.y.)`
}
# Specify the columns you want to modify
columns_to_modify <- paste0("d_", 1:100)
# Loop through the columns and randomly delete values
for (column in columns_to_modify) {
  non_na_indices <- which(!is.na(data_new1[, column]))
  # Check if there are enough non-NA values to delete
  if (length(non_na_indices) >= 10) {
    indices_to_delete <- sample(non_na_indices, 10)
    data_new1[indices_to_delete, column] <- NA
  } else {
    warning(paste("Not enough non-NA values in", column, "to delete."))
  }
}
print(data_new1)
# Specify the columns to be interpolated (d_1, d_2, ..., d_100)
columns_to_interpolate <- paste0("d_", 1:100)
# Loop through the columns and perform linear interpolation using na.approx
for (col in columns_to_interpolate) {
  data_new1[[col]] <- na.approx(data_new1[[col]], rule = 2)
}

write.csv(data_new1, "x.csv", row.names = FALSE)
# Columns to calculate RMSE for
columns_to_calculate_rmse <- paste0("d_", 1:100)
# Calculate RMSE for each column
rmse_values <- sapply(columns_to_calculate_rmse, function(col) {
  sqrt(mean((data_new1$`MAR TOC (gC/cm2/k.y.)` - data_new1[[col]])^2, na.rm = TRUE))
})
# Calculate the mean RMSE
mean_rmse <- mean(rmse_values)
# Print or use mean_rmse as needed
cat("Mean RMSE:", mean_rmse, "\n")
# Create a new row with the mean RMSE and NAs in other columns
new_row <- data.frame(
  `MAR TOC (gC/cm2/k.y.)` = mean_rmse,
  setNames(rep(NA_real_, ncol(data_new1) - 1), names(data_new1)[-1])
)
# Ensure the new row has the correct structure
new_row <- as.data.frame(matrix(NA, ncol = ncol(data_new1), nrow = 1))
names(new_row) <- names(data_new1)
# Assign the mean RMSE value
new_row$`MAR TOC (gC/cm2/k.y.)` <- mean_rmse
# Append the new row to the original dataframe
data_new1 <- rbind(data_new1, new_row)
# Print the mean RMSE
cat("Mean RMSE:", mean_rmse, "\n")
# Write the updated dataframe to CSV & xlsx
write.csv(data_new1, "x.csv", row.names = FALSE)
write_xlsx(data_new1, "output.xlsx")