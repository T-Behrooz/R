# Load required packages  
library(readxl)  
library(lpSolve)  
library(WriteXLS)  

# Read data from the Excel file  
df <- data.frame(read_excel(path = "d:/r.xlsx", sheet = "data"))  

# Define input and output parameters  
m <- 2  # number of inputs  
s <- ncol(df) - m  # number of outputs  
N <- nrow(df)  # number of DMUs  

# Initialize matrices for constraints  
f.con <- matrix(ncol = N, nrow = m + s + 1)  
inputs <- data.frame(df[, 1:m])  
outputs <- df[, (m + 1):(s + m)]  
f.obj <- c(rep(0, N), rep(1, s))  # Objective function for BCC  

# Initialize empty data frames to store results  
weights <- NULL  
effcrs <- NULL  

# Iterate through each DMU to run the BCC DEA model  
for (j in 1:N) {  
  # Right-hand side values for constraints  
  f.rhs <- c(as.numeric(df[j, 1:m]), as.numeric(df[j, (m + 1):(m + s)]), 1)  
  f.dir <- c(rep("=", m), rep(">=", s), "=")  # Directions for BCC model  
  
  # Construct the constraints  
  for (i in 1:m) {  
    f.con[i, 1:N] <- c(as.numeric(df[, i]))  # Input constraints  
  }  
  
  for (r in (m + 1):(s + m)) {  
    f.con[r, 1:N] <- c(as.numeric(df[, r]))  # Output constraints  
  }  
  
  # Set up the additional constraints for the BCC model  
  c1 <- rbind(matrix(0, m, m), matrix(0, s, m), 0)  
  c2 <- rbind(matrix(0, m, s), -diag(s), 0)  
  f.con[m + s + 1, 1:N] <- c(rep(1, N))  
  
  # Combine constraints for the linear programming problem  
  final.con <- cbind(f.con, c1, c2)  
  
  # Solve the linear programming problem using lp function  
  results <- lp("max", f.obj, final.con, f.dir, f.rhs, compute.sens = TRUE)  
  
  # Store the results  
  if (is.null(weights)) {  
    weights <- results$solution[seq(1:(N + s))]  
    effcrs <- results$objval  
  } else {  
    weights <- rbind(weights, results$solution[seq(1:(N + s))])  
    effcrs <- rbind(effcrs, results$objval)  
  }  
}  

# Create a DataFrame for output  
EEs <- data.frame(effcrs, weights)  
colnames(EEs) <- c('Efficiency', rep("lambda", N), rep('s-', m), rep('s+', s))  

# Write results to Excel file  
WriteXLS(EEs, "d:/r-bcc.xls", row.names = TRUE, col.names = TRUE)  

print("Results written to e:/output-bcc.xls")