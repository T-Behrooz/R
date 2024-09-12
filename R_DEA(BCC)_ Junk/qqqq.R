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

# BCC Input-Oriented Model  
theta <- numeric(n)  
s_in <- matrix(0, nrow = n, ncol = ncol(X))  
s_out <- matrix(0, nrow = n, ncol = ncol(Y))  
lambda <- matrix(0, nrow = n, ncol = n)  

# Start timer  
start_time <- Sys.time()  

for (i in 1:n) {  
  # Solve the LP problem  
  library(lpSolve)  
  obj <- c(rep(0, ncol(X)), rep(1, ncol(Y)), 1)  
  constr <- rbind(cbind(X, -Y, diag(n)), c(rep(0, ncol(X) + ncol(Y)), rep(1, n)))  
  dir <- c(rep("<=", ncol(X) + ncol(Y)), rep("==", n))  
  rhs <- c(rep(0, ncol(X) + ncol(Y)), 1)  
  res <- lp("min", obj, constr, dir, rhs, all.int = FALSE)  
  
  # Extract the results  
  theta[i] <- res$objval  
  s_in[i, ] <- pmax(0, res$solution[1:ncol(X)])  
  s_out[i, ] <- pmax(0, res$solution[(ncol(X)+1):(ncol(X)+ncol(Y))])  
  lambda[i, ] <- pmax(0, res$solution[(ncol(X)+ncol(Y)+1):(ncol(X)+ncol(Y)+n)])  
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
for (j in 1:ncol(s_in)) {  
  results_df[, paste0("Input_Slack_", j)] <- s_in[, j]  
}  

# Add separate columns for each output slack  
for (j in 1:ncol(s_out)) {  
  results_df[, paste0("Output_Slack_", j)] <- s_out[, j]  
}  

# Add columns for lambda values  
for (j in 1:n) {  
  results_df[, paste0("Lambda", j)] <- lambda[, j]  
}  

# Write results to Excel  
write_xlsx(results_df, "D:/res.xlsx")  

print(paste("Execution time:", execution_time))