library(deaR)  
library(readxl) 
library(writexl)  
library(dplyr)

# Load data
input_file <- "d:/11.xlsx"  
output_dir <- "d:/New_DEA_Results/" 
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

raw_data <- read_excel(input_file)

# Function to normalize data (min-max normalization)  
normalize_data <- function(x, type = "input") {  
  if (type == "input") {  
    (x - min(x, na.rm = TRUE)) / (max(x, na.rm = TRUE) - min(x, na.rm = TRUE))  
  } else if (type == "output") {  
    (max(x, na.rm = TRUE) - x) / (max(x, na.rm = TRUE) - min(x, na.rm = TRUE))  
  }  
}  

# Normalize inputs and outputs separately
normalized_data <- raw_data %>%
  mutate(
    Input1 = normalize_data(Input1, type = "input"),
    Input2 = normalize_data(Input2, type = "input"),
    Output1 = normalize_data(Output1, type = "output"),
    Output2 = normalize_data(Output2, type = "output")
  )

# Create DEA data structures
dea_data <- make_deadata(  
  raw_data,  
  dmus = 1,       # DMU column  
  inputs = 2:3,   # Input columns (Input1, Input2)  
  outputs = 4:5   # Output columns (Output1, Output2)  
)

dea_data_normalized <- make_deadata(  
  normalized_data,  
  dmus = 1,       # DMU column  
  inputs = 2:3,   # Input columns (Input1, Input2)  
  outputs = 4:5   # Output columns (Output1, Output2)  
)

# Calculate BCC (VRS) efficiency
result <- model_basic(dea_data, orientation = "io", rts = "vrs")
filename <- "None_Normalized_BCC_Result_(io).xlsx"
full_path <- file.path(output_dir, filename)
summary(result, exportExcel = TRUE, filename = full_path)

result <- model_basic(dea_data_normalized, orientation = "io", rts = "vrs")
filename <- "Normalized_BCC_Result_(io).xlsx"
full_path <- file.path(output_dir, filename)
summary(result, exportExcel = TRUE, filename = full_path)

# Calculate Sexton efficiency (additive DEA model)
# For unnormalized data
result_sexon_unnormalized <- cross_efficiency(dea_data, orientation = "io",selfapp =TRUE )
filename <- "None_Normalized_Sexton_Efficiency.xlsx"
full_path <- file.path(output_dir, filename)
summary(result_sexon_unnormalized, exportExcel = TRUE, filename = full_path)

# For normalized data
result_sexon_normalized <- cross_efficiency(dea_data_normalized, orientation = "io",selfapp =TRUE )
filename <- "Normalized_Sexton_Efficiency.xlsx"
full_path <- file.path(output_dir, filename)
summary(result_sexon_normalized, exportExcel = TRUE, filename = full_path)