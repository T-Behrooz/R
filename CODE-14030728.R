library(deaR)  
library(readxl)  

path_in = "C:/Users/Tizchang/Desktop/dea/Dataset_5_2.xlsx"  
path_CCR = paste0("C:/Users/Tizchang/Desktop/dea/result_CCR_", Sys.Date(), "_", format(Sys.time(), "%H-%M-%S"), ".xlsx")  
path_BCC = paste0("C:/Users/Tizchang/Desktop/dea/result_BCC_", Sys.Date(), "_", format(Sys.time(), "%H-%M-%S"), ".xlsx")  
path_vrs_malmquist = paste0("C:/Users/Tizchang/Desktop/dea/vrs_malmquist_", Sys.Date(), "_", format(Sys.time(), "%H-%M-%S"), ".xlsx")  
path_crs_malmquist = paste0("C:/Users/Tizchang/Desktop/dea/crs_malmquist_", Sys.Date(), "_", format(Sys.time(), "%H-%M-%S"), ".xlsx")  
path_vrs_ranking = paste0("C:/Users/Tizchang/Desktop/dea/vrs_ranking_", Sys.Date(), "_", format(Sys.time(), "%H-%M-%S"), ".xlsx")  

# Importing Data from Excel  
maldataset = read_excel(path_in, sheet = "mal")  
effdataset = read_excel(path_in, sheet = "data2")  

# Convert input and output variables to numeric  
maldataset$Input1 <- as.numeric(maldataset$Input1)  
maldataset$Input2 <- as.numeric(maldataset$Input2)  
maldataset$Output1 <- as.numeric(maldataset$Output1)  
maldataset$Output2 <- as.numeric(maldataset$Output2)   

# Envelopment Efficiency (BCC & CCR)  
Mal = make_deadata(effdataset, ni = 2, no = 2)  

result_model_bcc = model_basic(Mal, orientation = "io", rts = "vrs")  
summary(result_model_bcc, exportExcel = TRUE, filename = path_BCC, sheet_name = "BCC")  

result_model_ccr = model_basic(Mal, orientation = "io", rts = "crs")  
summary(result_model_ccr, exportExcel = TRUE, filename = path_CCR, sheet_name = "CCR")  
# Multiplier DEA model Efficiency (BCC & CCR) 

# Malmquist  
dataset = make_malmquist(maldataset, percol = 2, arrangement = "vertical", ni = 2, no = 2)  

result_vrs = malmquist_index(dataset, orientation = "oo", rts = "vrs", tc_vrs = FALSE)  
summary(result_vrs, exportExcel = TRUE, filename = path_vrs_malmquist)  

result_crs = malmquist_index(dataset, orientation = "oo", rts = "crs", tc_vrs = FALSE)  
summary(result_crs, exportExcel = TRUE, filename = path_crs_malmquist)  

# Sexton  
ranking = make_deadata(datadea = effdataset,   
                       inputs = 2:3,  
                       outputs = 4:5)  

result_sexton = cross_efficiency(ranking,  
                                 dmu_eval = NULL,  
                                 dmu_ref = NULL,  
                                 epsilon = 0,  
                                 orientation = "oo",  
                                 rts = "vrs",  
                                 selfapp = TRUE)  

summary(result_sexton, exportExcel = TRUE, filename = path_vrs_ranking)