# Load required packages  
library(readxl)  
library(lpSolve)  
library(rJava)  
library(WriteXLS)  

# Read data from Excel file  
df <- data.frame(read_excel(path = "d:/r.xlsx", sheet = "data"))  

# Define input and output parameters dynamically based on the dataframe  
m <- 2  # Number of inputs (I)  
s <- ncol(df) - m  # Number of outputs (O)  
N <- nrow(df)  # Number of DMUs  

# Initialize matrices for constraints  
f.con <- matrix(ncol = N + 1, nrow = m + s + 1)  

# Iterate through each DMU to run analysis  
for (j in 1:N) {  
  f.rhs <- c(rep(0, m), unlist(unname(df[j, (m + 1):(m + s)])), 1)  
  f.dir <- c(rep("<=", m), rep(">=", s), "=")  
  f.obj <- c(1, rep(0, N))  
  
  # Construct constraints for inputs  
  for (i in 1:m) {  
    f.con[i, 1:(N + 1)] <- c(as.numeric(-df[j, i]), df[j, i])  
  }  
  
  # Construct constraints for outputs  
  for (r in (m + 1):(s + m)) {  
    f.con[r, 1:(N + 1)] <- c(0, as.numeric(df[, r]))  
  }  
  
  # Add overall constraint  
  f.con[m + s + 1, 1:(N + 1)] <- c(0, rep(1, N))  
  
  # Solve the linear programming problem  
  results <- lp("min", as.numeric(f.obj), f.con, f.dir, f.rhs, scale = 0, compute.sens = TRUE)  
  
  # Store results  
  if (j == 1) {  
    weights <- results$solution[1]  
    eff <- results$objval  
    lambdas <- results$duals[seq(2, N + 1)]  
  } else {  
    weights <- rbind(weights, results$solution[1])  
    eff <- rbind(eff, results$objval)  
    lambdas <- rbind(lambdas, results$solution[seq(2, N + 1)])  
  }  
}  

# Prepare to run second phase of DEA  
Es <- data.frame(eff)  
theta_0 <- unname(Es[, 1])  
f.con1 <- matrix(ncol = N + m, nrow = m + s + 1)  
f.obj1 <- c(rep(0, N), rep(1, m), rep(1, s))  
f.dir1 <- c(rep("=", m + s + 1))  
text <- matrix(ncol = m, nrow = N)  

# Iterate through each input  
for (i in 1:m) {  
  text[i, ] <- theta_0 * df[, i]  
  for (j in 1:N) {  
    # Construct input constraints  
    for (i in 1:m) {  
      f.con1[i, 1:N] <- c(as.numeric(df[, i]))  
    }  
    # Construct output constraints  
    for (r in (m + 1):(s + m)) {  
      f.con1[r, 1:N] <- c(as.numeric(df[, r]))  
    }  
    # Add overall constraint  
    f.con1[m + s + 1, 1:N] <- c(rep(1, N))  
    f.con1[(N + 1):(N + m)] <- rbind(-diag(m), matrix(0, m, 0))  
    f.con1[(N + m + 1):(N + m + 1)] <- rbind(matrix(0, s), diag(s), 0)  
    
    # Define right-hand side for the second phase  
    rhs1 <- c(unlist(unname(text[j, ])), unlist(unname(data.frame(df)[, (m + 1):(m + s)])), 1)  
    
    # Solve the linear programming problem for the second phase  
    results <- lp("max", as.numeric(f.obj1), f.con1, f.dir1, rhs1, scale = 0, compute.sens = TRUE)  
    
    # Store results for second phase  
    if (j == 1) {  
      weights1 <- results$solution  
      eff1 <- results$objval  
      lambdas1 <- results$duals[seq(1, N)]  
    } else {  
      weights1 <- rbind(weights1, results$solution)  
      eff1 <- rbind(eff1, results$objval)  
      lambdas1 <- rbind(lambdas1, results$duals[seq(1, N)])  
    }  
  }  
}  

# Combine results into a data frame for output  
Es1 <- data.frame(eff1, eff)  
colnames(Es1) <- c('Phase2', 'Phase1')  

# Write results to an Excel file  
WriteXLS(Es1, "d:/12-TWO-Phase-CCR.xls", row.names = TRUE, col.names = TRUE)  
print("Results written to d:/12-TWO-Phase-CCR.xls")