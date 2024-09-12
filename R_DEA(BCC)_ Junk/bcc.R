library(deaR)
library(readxl)
#-------------------------------------- envelopment eff(BCC& CCR)---------------
start.time <- Sys.time()
Mal_Data2 = read_excel("d:/data.xlsx",sheet = "data")
Mal = read_data(Mal_Data2,ni=3,no=5)
result_model= model_basic(Mal, orientation = "io" , rts =  "vrs")
summary(result_model, exportExcel = TRUE, filename = "d:/vrs-effi.xlsx")

################################################################################
# Mal_Data3 = read_excel("d:/bankdata.xlsx",sheet = "data")
# Mal2 = read_data(Mal_Data3,ni=2,no=1)
# result_model= model_basic(Mal2, orientation = "io" , rts =  "crs")
# summary(result_model, exportExcel = TRUE, filename = "d:/crs-effi.xlsx")
################################################################################
end.time <- Sys.time()
time.taken <- round(end.time - start.time,2)
time.taken
#-------------------------------------------------------------------------------