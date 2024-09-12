library(readxl)  
library(writexl)  

# Load data from Excel file  
data <- read_excel("D:/r.xlsx")  

# Identify column names  
input_cols <- grep("^I[0-9]+$", names(data), value = TRUE)  
output_cols <- grep("^O[0-9]+$", names(data), value = TRUE)  
dmu_col <- grep("^DMU$", names(data), value = TRUE)  

# Extract input and output data  
X <- as.matrix(data[, input_cols])  
Y <- as.matrix(data[, output_cols])  
DMU <- data[, dmu_col]  

# Number of DMUs  
n <- nrow(data)  
m <- ncol(X) # Number of inputs  
s <- ncol(Y) # Number of outputs  

# BCC Input-Oriented Model  
theta <- numeric(n)  
s_in <- matrix(0, nrow = n, ncol = m)  
s_out <- matrix(0, nrow = n, ncol = s)  
lambda <- matrix(0, nrow = n, ncol = n)  

# Start timer  
start_time <- Sys.time()  

for (i in 1:n) {  
  # Solve the LP problem  
  library(lpSolve)  
  obj <- c(rep(0, m), rep(1, s), -1)  
  constr <- rbind(cbind(X, -Y, diag(n)), c(rep(0, m + s), rep(1, n)))  
  dir <- c(rep("<=", m + s), rep("==", n))  
  rhs <- c(X[i,], -Y[i,], 1)  
  res <- lp("min", obj, constr, dir, rhs, all.int = FALSE)  
  
  # Extract the results  
  theta[i] <- res$objval  
  s_in[i, ] <- res$solution[1:m]  
  s_out[i, ] <- res$solution[(m+1):(m+s)]  
  lambda[i, ] <- res$solution[(m+s+1):(m+s+n)]  
}  

# Stop timer  
end_time <- Sys.time()  
execution_time <- end_time - start_time  

# Create a data frame with the results  
results_df <- data.frame(  
  DMU = DMU,  
  Efficiency = theta  
)  

# Add separate columns for each input slack  
for (j in 1:m) {  
  results_df[, paste0("Input_Slack_", j)] <- s_in[, j]  
}  

# Add separate columns for each output slack  
for (j in 1:s) {  
  results_df[, paste0("Output_Slack_", j)] <- s_out[, j]  
}  

# Add columns for lambda values  
for (j in 1:n) {  
  results_df[, paste0("Lambda", j)] <- lambda[, j]  
}  

# Write results to Excel  
write_xlsx(results_df, "D:/res.xlsx")  

print(paste("Execution time:", execution_time))