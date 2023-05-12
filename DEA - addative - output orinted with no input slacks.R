library(readxl)
library(lpSolve)
library(rJava)
library(WriteXLS)

df <- data.frame(read_excel(path = "e:/sample.xlsx", sheet = "data"))
m <- 2
s <- ncol(df) - m
N <- nrow(df)
f.con <- matrix(ncol = N, nrow = m + s + 1)
inputs <- data.frame(df[, 1:m])
outputs <- df[, (m + 1):(s + m)]
f.obj <- c(rep(0, N), rep(0, m), rep(1, s))

for (j in 1:N) {
  f.rhs <- c(as.numeric(df[j, 1:m]), as.numeric(df[j, (m + 1):(m + s)]), 1)
  f.dir <- c(rep("=", m), rep("=", s), "=")
  
  for (i in 1:m) {
    f.con[i, 1:N] <- c(as.numeric(df[, i]))
  }
  
  for (r in (m + 1):(s + m)) {
    f.con[r, 1:N] <- c(as.numeric(df[, r]))
  }
  
  c1 <- rbind(matrix(0,m,m), matrix(0, s, m), 0)
  c2 <- rbind(matrix(0, m, s), -diag(s), 0)
  f.con[m + s + 1, 1:N] <- c(rep(1, N))
  final.con <- cbind(f.con, c1, c2)
  results <- lp("max", as.numeric(f.obj), final.con, f.dir, f.rhs, compute.sens = TRUE)
  
  if (j == 1) {
    weights <- results$solution[seq(1:(N + m + s))]
    effcrs <- results$objval
  } else {
    weights <- rbind(weights, results$solution[seq(1:(N + m + s))])
    effcrs <- rbind(effcrs, results$objval)
  }
}

EEs <- data.frame(effcrs, weights)
colnames(EEs) <- c('sum', rep("lambda", N), rep('s-', m), rep('s+', m))
WriteXLS(EEs, "e:/output-additive.xls", row.names = TRUE, col.names = TRUE)