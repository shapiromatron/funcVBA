input_tbl <- read.table(file="C:/Program Files/IEUBKwin1_1 Build11/Output/PnCB-Age2.csv", header=TRUE, sep=",")
attach(input_tbl)
library(scatterplot3d)
scatterplot3d(GI_Uptake, Air_Uptake, PbB, pch=16, highlight.3d=TRUE, type="h", main="title")