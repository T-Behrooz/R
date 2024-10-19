library(deaR)
library(readxl)


path_in="C:/Users/Tizchang/Desktop/dea/Dataset_5_2.xlsx"
path_CCR="C:/Users/Tizchang/Desktop/dea/result_CCR.xlsx"
path_BCC="C:/Users/Tizchang/Desktop/dea/result_BCC.xlsx"

#----------------------------------------- Importing Data from Excel-----------------------------------

maldataset = read_excel(path_in,sheet = "mal")
effdataset = read_excel(path_in,sheet = "data2")

# Convert input and output variables to numeric  
maldataset$Input1 <- as.numeric(maldataset$Input1)  
maldataset$Input2 <- as.numeric(maldataset$Input2)  
maldataset$Output1 <- as.numeric(maldataset$Output1)  
maldataset$Output2 <- as.numeric(maldataset$Output2) 

#-------------------------------------- envelopment eff(BCC& CCR)---------------------------------------

Mal = make_deadata(effdataset,ni=2,no=2)

result_model= model_basic(Mal, orientation = "io" , rts =  "vrs")

summary(result_model, exportExcel = TRUE, filename = path_BCC,sheet_name="Bcc")

result_model= model_basic(Mal, orientation = "io" , rts =  "crs")

summary(result_model, exportExcel = TRUE, filename = path_CCR,,sheet_name="CCR")

#----------------------------------------Malmquist-------------------------------------------------------

dataset = make_malmquist(maldataset ,percol=2, arrangement="vertical",ni=2,no=2)

result_vrs = malmquist_index(dataset,orientation = "oo",rts =  "vrs",tc_vrs = FALSE)

result_crs = malmquist_index(dataset,orientation = "oo",rts =  "crs",tc_vrs = FALSE)

#---------------------------------------------------------------------------------------------------------
ranking = make_deadata(datadea = effdataset, 
                     inputs = 2:4, 
                     outputs = 5:6)
result_sexton = cross_efficiency(ranking,
                 dmu_eval = NULL,
                 dmu_ref = NULL,
                 epsilon = 0,
                 orientation ="oo",
                 rts ="vrs",
                 selfapp = TRUE
                )
#---------------------------------------------------------------------------------------------------------

summary(result_vrs, exportExcel = TRUE, filename = "C:/Users/Tizchang/Desktop/dea/vrs_malmquist.xlsx")
summary(result_crs, exportExcel = TRUE, filename = "C:/Users/Tizchang/Desktop/dea/crs_malmquist.xlsx")
summary(result_sexton, exportExcel = TRUE, filename = "C:/Users/Tizchang/Desktop/dea/vrs_ranking.xlsx")
#---------------------------------------------------------------------------------------------------------

