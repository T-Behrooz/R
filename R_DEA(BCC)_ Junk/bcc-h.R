library(readxl)  
library(lpSolve)  
library(rJava)  
library(WriteXLS)  
df=data.frame(read_excel(path = "d:/r.xlsx", sheet = "1"))  
m=2  
s=ncol(df)-m  
N=nrow(df)  
f.con=matrix(ncol=N+1,nrow=m+s+1)  
for (j in 1:N) {  
  f.rhs = c(rep(0,m),unlist(unname(df[j,(m+1):(m+s)])), 1)  
  f.dir = c(rep("<=",m),rep(">=",s), "=")  
  f.obj = c(1, rep(0,N))  
  
  for(i in 1:m)  
  {  
    f.con[i,1:(N+1)]=c(as.numeric(-df[j,i]),df[j,i])  
  }  
  
  for(r in (m+1):(s+m))  
  {  
    f.con[r,1:(N+1)]=c(0,as.numeric(df[,r]))  
  }  
  
  f.con[m+s+1, 1:(N+1)]=c(0, rep(1,N))  
  results = lp("min", as.numeric(f.obj), f.con, f.dir, f.rhs, scale=0, compute.sens=TRUE)  
  
  if (j==1) {  
    weights = results$solution[1]  
    eff = results$objval  
    lambdas = results$duals[seq(2,N+1)]  
  } else {  
    weights = rbind(weights, results$solution[1])  
    eff = rbind(eff, results$objval)  
    lambdas = rbind(lambdas, results$solution[seq(2,N+1)])  
  }  
}  
Es = data.frame(eff)  
theta_0 = unname(Es[,1])  
f.con=matrix(ncol=N+m,s=nrow=m+s+1)  
f.obj = c(rep(0,N),rep(1,m),rep(1,s))
f.dir = c(rep("=", m+s+1))  
text = matrix(ncol=m,nrow=N)  
for(i in 1:m)  
{  
  text[i,] = eta_0*df[,i]  
  for(j in 1:N) {  
    for(i in 1:m) {  
      f.con1[i, 1:N] = c(as.numeric(df[,i]))  
    }  
    for(r in (m+1):(s+m)) {  
      f.con1[r, L:N] = c(as.numeric(df[,r]))  
    }  
    f.con1[m+s+1, 1:N] = c(rep(1,N))  
    f.con1[(N+1):(N+m)]= rbind(-diag(m), matrix(0,m,0))  
    f.con1[(N+m+1):(N+m+1)] = rbind(matrix(0,s), diag(s),0)  
    rhs1 = c(unlist(unname(text[j,])), unlist(unname(data.frame(df)[,(m+1):(m+s)])), 1)  
    results = lp("max", as.numeric(f.obj), f.con1, f.dir, rhs1, scale=0, compute.sens=TRUE)  
    if(j==1) {  
      weights1 = results$solution  
      eff1 = results$objval  
      lambdas = results$duals[seq(1,N)]  
    } else {  
      weights1 = rbind(weights1, results$solution)  
      eff1 = rbind(eff1, results$objval)  
      lambdas = rbind(lambdas, results$duals[seq(1,N)])  
    }  
  }  
}  
Es1 = data.frame(eff1, eff)  
colnames(Es1) = c('phase2', 'phase1')  
WriteXLS(Es1, "12-TWO-Phase-CCR.xls", row.names = TRUE, col.names = TRUE)