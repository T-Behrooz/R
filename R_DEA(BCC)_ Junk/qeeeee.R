# Load required packages  
library(readxl)  
library(lpSolve)  
library(openxlsx)  

# Read data from Excel file  
data <- read_excel("d:\\r.xlsx")  

# Extract input and output variables  
inputs <- data[, grepl("^I[0-9]+", names(data))]  
outputs <- data[, grepl("^o[0-9]+", names(data))]  
dmus <- data$dmu  

# Solve BCC model for each DMU  
theta <- numeric(nrow(data))  
slacks_in <- matrix(nrow = nrow(data), ncol = ncol(inputs))  
slacks_out <- matrix(nrow = nrow(data), ncol = ncol(outputs))  
lambdas <- matrix(nrow = nrow(data), ncol = nrow(data))  

for (i in 1:nrow(data)) {  
  x <- as.matrix(inputs[i, ])  
  y <- as.matrix(outputs[i, ])  
  
  # Solve BCC model  
  model <- lp(  
    direction = "min",  
    objective.in = c(rep(1, ncol(x)), rep(0, ncol(y))),  
    const.mat = cbind(x, -y, diag(nrow(x)), matrix(0, nrow(x), ncol(y))),  
    const.dir = c(rep("<=", nrow(x)), rep(">=", ncol(y)), rep("=", 1)),  
    const.rhs = c(rep(0, nrow(x)), y, 1)  
  )  
  
  if (model$status == 0) {  
    theta[i] <- model$objval  
    slacks_in[i, ] <- model$solution[1:ncol(x)]  
    slacks_out[i, ] <- model$solution[(ncol(x)+1):(ncol(x)+ncol(y))]  
    lambdas[i, ] <- model$solution[(ncol(x)+ncol(y)+1):(ncol(x)+ncol(y)+nrow(data))]  
  } else {  
    # Handle cases where the model doesn't solve successfully  
    theta[i] <- NA  
    slacks_in[i, ] <- rep(NA, ncol(x))  
    slacks_out[i, ] <- rep(NA, ncol(y))  
    lambdas[i, ] <- rep(NA, nrow(data))  
  }  
}  

# Write results to Excel file  
write.xlsx(data.frame(dmu = dmus,   
                      Efficiency = theta,   
                      Slack_in = slacks_in,  
                      Slack_out = slacks_out,  
                      Lambda = lambdas),   
           "d:\\r_result.xlsx",   
           sheetName = "Result",   
           row.names = FALSE)