# Load required packages  
library(readxl)  
library(lpSolve)  
library(WriteXLS)  

# Read data from Excel file  
df <- data.frame(read_excel(path = "d:/r.xlsx", sheet = "data"))  

# Validate and convert data to numeric  
df[] <- lapply(df, as.numeric)  # Ensure all columns are numeric  

# Handle NA values  
df <- na.omit(df)  # Remove rows with NA values  
# OR alternatively replace NA values with the mean of the columns  
# for (i in 1:ncol(df)) {  
#   df[is.na(df[, i]), i] <- mean(df[, i], na.rm = TRUE)  # Replace NA with mean  
# }  

# Ensure no NA values left  
if (anyNA(df)) {  
  stop("There are still NA values in the data.")  
}  

# Define input and output parameters  
m <- 2  # Number of inputs (I)  
s <- ncol(df) - m  # Number of outputs (O)  
N <- nrow(df)  # Number of DMUs  

# Initialize matrices for constraints  
f.con <- matrix(0, nrow = m + s + 1, ncol = N + 1)  
f.rhs <- matrix(0, nrow = m + s + 1, ncol = 1)  
f.dir <- character(m + s + 1)  
f.obj <- c(1, rep(0, N))  

# Iterate through each DMU to run analysis  
for (j in 1:N) {  
  # Construct constraints for inputs  
  for (i in 1:m) {  
    f.con[i, j + 1] <- -df[j, i]  
  }  
  
  # Construct constraints for outputs  
  for (r in (m + 1):(s + m)) {  
    f.con[r, j + 1] <- df[j, r]  
  }  
  
  # Add overall constraint  
  f.con[m + s + 1, j + 1] <- 1  
  
  # Set right-hand side values  
  f.rhs[1:m, 1] <- 0  
  f.rhs[(m + 1):(m + s), 1] <- unlist(unname(df[j, (m + 1):(m + s)]))  
  f.rhs[m + s + 1, 1] <- 1  
  
  # Set constraint directions  
  f.dir[1:m] <- "<="  
  f.dir[(m + 1):(m + s)] <- ">="  
  f.dir[m + s + 1] <- "="  
  
  # Solve linear programming problem  
  results <- lp("min", as.numeric(f.obj), f.con, f.dir, f.rhs, scale = 0, compute.sens = TRUE)  
  
  # Store results  
  if (j == 1) {  
    weights <- results$solution[1]  
    eff <- results$objval  
    lambdas <- results$duals[seq(2, N + 1)]  
  } else {  
    weights <- rbind(weights, results$solution[1])  
    eff <- rbind(eff, results$objval)  
    lambdas <- rbind(lambdas, results$duals[seq(2, N + 1)])  
  }  
}  

# Prepare to run the second phase of DEA  
Es <- data.frame(eff)  
theta_0 <- unname(Es[, 1])  
f.con1 <- matrix(ncol = N + m, nrow = m + s + 1)  
f.obj1 <- c(rep(0, N), rep(1, m), rep(1, s))  
f.dir1 <- c(rep("=", m + s + 1))  
text <- matrix(ncol = m, nrow = N)  

# Iterate through each input  
for (k in 1:N) {  
  text[k, ] <- theta_0[k] * df[k, 1:m]  
}  

# Construct constraints for second phase  
for (j in 1:N) {  
  for (i in 1:m) {  
    f.con1[i, 1:N] <- as.numeric(df[, i])  
  }  
  for (r in (m + 1):(s + m)) {  
    f.con1[r, 1:N] <- as.numeric(df[, r])  
  }  
  f.con1[m + s + 1, 1:N] <- rep(1, N)  
  f.con1[(N + 1):(N + m), 1:N] <- -diag(m)  
  f.con1[(N + m + 1):(N + m + 1), 1:N] <- diag(s)  
  f.con1[nrow(f.con1), N + 1] <- 1  
  
  rhs1 <- c(text[j, ], unlist(unname(df[j, (m + 1):(m + s)])), 1)  
  
  # Solve linear programming problem for the second phase  
  results <- lp("max", as.numeric(f.obj1), f.con1, f.dir1, rhs1, scale = 0, compute.sens = TRUE)  
  
  # Store results for second phase  
  if (j == 1) {  
    second_phase_weights <- results$solution[1:N]  
    second_phase_eff <- results$objval  
    second_phase_lambdas <- results$solution[(N + 1):(N + m + s)]  
  } else {  
    second_phase_weights <- rbind(second_phase_weights, results$solution[1:N])  
    second_phase_eff <- rbind(second_phase_eff, results$objval)  
    second_phase_lambdas <- rbind(second_phase_lambdas, results$solution[(N + 1):(N + m + s)])  
  }  
}  

# Combine results  
results_df <- data.frame(  
  DMU = 1:N,  
  Efficiency = eff,  
  Weights = weights,  
  Lambdas = apply(lambdas, 1, paste, collapse = ", "),  
  Second_Phase_Efficiency = second_phase_eff,  
  Second_Phase_Weights = apply(second_phase_weights, 1, paste, collapse = ", "),  
  Second_Phase_Lambdas = apply(second_phase_lambdas, 1, paste, collapse = ", ")  
)  

# Write results to Excel file  
WriteXLS("results.xlsx", results_df)