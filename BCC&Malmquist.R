library(deaR)
library(readxl)
library(writexl)
library(dplyr)
library(ggplot2)
library(gridExtra)

# Set input file path and output directory
input_file <- "d:/11.xlsx"
output_dir <- "d:/BCC_&_MALMQUIST_DEA_Results/"
dir.create(output_dir, showWarnings = FALSE, recursive = TRUE)

# Read raw data
raw_data <- read_excel(input_file)

# Read data for Malmquist analysis
mal_raw_data <- read_excel(input_file, sheet = "Mal")

# Function to normalize data (min-max normalization)
normalize_data <- function(x, type = "input") {
  # Handle NA values
  x <- x[!is.na(x)]
  
  if (type == "input") {
    # For inputs (minimization), normalize to [0,1]
    (x - min(x)) / (max(x) - min(x))
  } else if (type == "output") {
    # For outputs (maximization), normalize to [0,1]
    (max(x) - x) / (max(x) - min(x))
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

# Prepare DEA data
dea_data <- make_deadata(
  raw_data,
  dmus = 1,       # DMU column
  inputs = 2:3,   # Input columns (Input1, Input2)
  outputs = 4:5   # Output columns (Output1, Output2)
)

# Prepare normalized DEA data
dea_data_normalized <- make_deadata(
  normalized_data,
  dmus = 1,       # DMU column
  inputs = 2:3,   # Input columns (Input1, Input2)
  outputs = 4:5   # Output columns (Output1, Output2)
)

# Perform BCC model on non-normalized data
result <- model_basic(dea_data, orientation = "io", rts = "vrs")

# Export results
filename <- "None_Normalized_BCC_Result_(io).xlsx"
full_path <- file.path(output_dir, filename)
write_xlsx(summary(result), full_path)

# Perform BCC model on normalized data
result_normalized <- model_basic(dea_data_normalized, orientation = "io", rts = "vrs")

# Export normalized results
filename <- "Normalized_BCC_Result_(io).xlsx"
full_path <- file.path(output_dir, filename)
write_xlsx(summary(result_normalized), full_path)

# Prepare Malmquist data
data_example2 <- make_malmquist(
  mal_raw_data,
  percol = 2,
  arrangement = "vertical",
  ni = 2,
  no = 2
)

# Calculate Malmquist index
Mal_result <- malmquist_index(data_example2, orientation = "io")

# Export Malmquist results
filename <- "Malmquist_BCC_Result_(io).xlsx"
full_path <- file.path(output_dir, filename)
write_xlsx(summary(Mal_result), full_path)

plot(Mal_result)