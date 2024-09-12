library(readxl)  
library(writexl)  

# Load data from Excel file  
data <- read_excel("d:/databank.xlsx")  

# Extract input and output data  
X <- data[, 1:3]  # 3 inputs  
Y <- data[, 4:8]  # 5 outputs  

# Number of DMUs  
n <- nrow(data)  

# Initialize vectors for results  
theta <- rep(0, n)  # Efficiency scores  
slack_in <- matrix(0, n, 3)  # Input slacks  
slack_out <- matrix(0, n, 5)  # Output slacks  
lambda <- matrix(0, n, n)  # Reference set  

# Start timer  
start_time <- Sys.time()  

# BCC Input-Oriented Model  
for (i in 1:n) {  
  # Solve the LP problem  
  library(lpSolve)  
  obj <- c(rep(1, 3), rep(0, n))  
  constr <- cbind(X, diag(n))  
  dir <- c(rep("<=", 3), rep(">=", n))  
  rhs <- c(X[i, ], rep(1, n))  
  res <- lp("min", obj, constr, dir, rhs, all.int = FALSE)  
  
  # Extract the results  
  theta[i] <- res$objval  
  slack_in[i, ] <- res$solution[1:3]  
  slack_out[i, ] <- res$solution[4:(3+5)]  
  lambda[i, ] <- res$solution[(4+5):length(res$solution)]  
}  

# Stop timer  
end_time <- Sys.time()  
execution_time <- end_time - start_time  

# Write results to Excel  
write_xlsx(list(Efficiency = theta,   
                Input_Slack = slack_in,  
                Output_Slack = slack_out,  
                Lambda = lambda),   
           "results.xlsx")  

print(paste("Execution time:", execution_time))