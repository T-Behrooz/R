# Load required packages  
library(readxl)  
library(lpSolve)  
library(openxlsx)  

# Read data from Excel file  
data <- read_excel("d:\\r.xlsx")  

# Extract input and output variables  
inputs <- as.matrix(data[, grepl("^I[0-9]+", names(data))])  
outputs <- as.matrix(data[, grepl("^O[0-9]+", names(data))])  

# Check the column names in the data dataframe  
print(names(data))  

# Assign the appropriate column name for DMUs  
dmus <- data[[1]]  

# Prepare matrices for results  
theta <- numeric(nrow(data))  
slacks_in <- matrix(NA, nrow = nrow(data), ncol = ncol(inputs))  
slacks_out <- matrix(NA, nrow = nrow(data), ncol = ncol(outputs))  
lambdas <- matrix(NA, nrow = nrow(data), ncol = nrow(data))  
computation_time <- numeric(nrow(data))  # To store computation times  

# Solve BCC model for each DMU  
for (i in 1:nrow(data)) {  
  x <- inputs[i, , drop = FALSE]  
  y <- outputs[i, , drop = FALSE]  
  
  # Debugging output  
  print(paste("Processing DMU:", dmus[i]))  
  print("Input Matrix (x):")  
  print(x)  
  print("Output Matrix (y):")  
  print(y)  
  
  # Convert y to a vector  
  y <- as.vector(y)  
  
  # Start timer  
  start_time <- Sys.time()  
  
  # Ensure dimensions match for the constraints  
  model <- lp(  
    direction = "min",  
    objective.in = c(rep(1, ncol(x)), rep(0, length(y))),  
    const.mat = rbind(x, -y, diag(ncol(x))),  
    const.dir = c(rep("<=", nrow(x)), rep(">=", length(y)), "="),  
    const.rhs = c(rep(0, nrow(x)), y, 1)  
  )  
  
  # End timer  
  end_time <- Sys.time()  
  computation_time[i] <- as.numeric(difftime(end_time, start_time, units = "secs"))  # Compute the time taken  
  
  print(model$status)  # Check status of the optimization  
  
  if (model$status == 0) {  
    theta[i] <- model$objval  
    slacks_in[i, ] <- model$solution[1:ncol(x)]  
    slacks_out[i, ] <- model$solution[(ncol(x)+1):(ncol(x)+length(y))]  
    lambdas[i, ] <- model$solution[(ncol(x)+length(y)+1):(ncol(x)+length(y)+nrow(data))]  
  } else {  
    theta[i] <- NA  
    slacks_in[i, ] <- rep(NA, ncol(x))  
    slacks_out[i, ] <- rep(NA, length(y))  
    lambdas[i, ] <- rep(NA, nrow(data))  
  }  
  
  # Print warning messages if exists  
  if (model$status != 0) {  
    print(paste("Warning: Model failed for DMU:", dmus[i]))  
  }  
}  

# Check for any warnings  
if(length(warnings()) > 0) {  
  print("Warnings encountered:")  
  print(warnings())  
} else {  
  print("No warnings were encountered.")  
}  

# Write results to Excel file  
write.xlsx(data.frame(dmu = dmus,   
                      Efficiency = theta,   
                      Slack_in = slacks_in,  
                      Slack_out = slacks_out,  
                      Lambda = lambdas,  
                      Computation_Time = computation_time),   
           "d:\\r_result.xlsx",   
           sheetName = "Result",   
           rowNames = FALSE)  # Corrected argument