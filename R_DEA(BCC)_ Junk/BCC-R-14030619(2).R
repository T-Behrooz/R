library(readxl)  
library(writexl)  

# Load data from Excel file  
data <- read_excel("D:/bankdata.xlsx")  

# Extract input and output data  
X <- data[2:(nrow(data)), 2:4]  # Inputs B, C, D  
Y <- data[2:(nrow(data)), 5:9]  # Outputs E, F, G, H, I  

# Number of DMUs  
n <- nrow(data) - 1  # Exclude header row  

# Initialize result lists  
theta_list <- list()  
slack_in_list <- list()  
slack_out_list <- list()  
lambda_list <- list()  

# Start timer  
start_time <- Sys.time()  

# BCC Input-Oriented Model  
for (i in 1:n) {  
  # Solve the LP problem  
  library(lpSolve)  
  obj <- c(rep(1, 3), rep(0, n))  
  constr <- cbind(X[i, ], diag(n))  
  dir <- c(rep("<=", 3), rep(">=", n))  
  rhs <- c(X[i, ], rep(1, n))  
  res <- lp("min", obj, constr, dir, rhs, all.int = FALSE)  
  
  # Extract the results  
  theta_list[[i]] <- res$objval  
  slack_in_list[[i]] <- res$solution[1:3]  
  slack_out_list[[i]] <- res$solution[4:(3+5)]  
  lambda_list[[i]] <- res$solution[(4+5):length(res$solution)]  
}  

# Stop timer  
end_time <- Sys.time()  
execution_time <- end_time - start_time  

# Convert lists to matrices  
theta <- unlist(theta_list)  
slack_in <- do.call(rbind, slack_in_list)  
slack_out <- do.call(rbind, slack_out_list)  
lambda <- do.call(rbind, lambda_list)  

# Write results to Excel  
write_xlsx(list(DMU = data$A[2:(n+1)],   
                Efficiency = theta,   
                Input_Slack = slack_in,  
                Output_Slack = slack_out,  
                Lambda = lambda),   
           "D:/results.xlsx")  

print(paste("Execution time:", execution_time))
