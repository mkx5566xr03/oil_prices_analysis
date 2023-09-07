install.packages("openxlsx")

library(openxlsx)

#讀取資料
df = read.xlsx("C:/Users/howger/Desktop/Time Series/Multivariate Time Series Analysis  With R and Financial Applications/資料庫價格20221101_v6-石化_(002).xlsx" , sheet = "原物料價格",startRow = 3 , detectDates = TRUE)

head(df)

rm()

#做9個變數之走勢圖
par(mfrow = c(3,3))
x1 = df[,2]
x2 = df[,3]
x3 = df[,4]
x4 = df[,5]
x5 = df[,6]
x6 = df[,7]
x7 = df[,8]
x8 = df[,9]
x9 = df[,10]
ts.plot(x1,ylab="",xlab="",main="原油布蘭特北海現貨價(美元/桶)")
ts.plot(x2,ylab="",xlab="",main="布蘭特原油倫敦ICE近月期貨收盤價(美元/桶)")
ts.plot(x3,ylab="",xlab="",main="輕油(石油腦)日本C＆F(美元/公噸)")
ts.plot(x4,ylab="",xlab="",main="乙烯遠東區CFR(美元/公噸)")
ts.plot(x5,ylab="",xlab="",main="丙烯東南亞CFR(美元/公噸)")
ts.plot(x6,ylab="",xlab="",main="苯東南亞CFR(美元/公噸)")
ts.plot(x7,ylab="",xlab="",main="低密度聚乙烯LDPE通用級遠東區CFR(美元/公噸)")
ts.plot(x8,ylab="",xlab="",main="高密度聚乙烯HDPE射出級遠東區CFR(美元/公噸)")
ts.plot(x9,ylab="",xlab="",main="聚丙烯PP吹袋級遠東區CFR(美元/公噸)")

#嘗試取log變數變換
lx1 = log(df[,2])
lx2 = log(df[,3])
lx3 = log(df[,4])
lx4 = log(df[,5])
lx5 = log(df[,6])
lx6 = log(df[,7])
lx7 = log(df[,8])
lx8 = log(df[,9])
lx9 = log(df[,10])
ts.plot(lx1,ylab="",xlab="",main="原油布蘭特北海現貨價(美元/桶)")
ts.plot(lx2,ylab="",xlab="",main="布蘭特原油倫敦ICE近月期貨收盤價(美元/桶)")
ts.plot(lx3,ylab="",xlab="",main="輕油(石油腦)日本C＆F(美元/公噸)")
ts.plot(lx4,ylab="",xlab="",main="乙烯遠東區CFR(美元/公噸)")
ts.plot(lx5,ylab="",xlab="",main="丙烯東南亞CFR(美元/公噸)")
ts.plot(lx6,ylab="",xlab="",main="苯東南亞CFR(美元/公噸)")
ts.plot(lx7,ylab="",xlab="",main="低密度聚乙烯LDPE通用級遠東區CFR(美元/公噸)")
ts.plot(lx8,ylab="",xlab="",main="高密度聚乙烯HDPE射出級遠東區CFR(美元/公噸)")
ts.plot(lx9,ylab="",xlab="",main="聚丙烯PP吹袋級遠東區CFR(美元/公噸)")

#嘗試取diff log 變數變換

dlx1 = diff(log(df[,2]))
dlx2 = diff(log(df[,3]))
dlx3 = diff(log(df[,4]))
dlx4 = diff(log(df[,5]))
dlx5 = diff(log(df[,6]))
dlx6 = diff(log(df[,7]))
dlx7 = diff(log(df[,8]))
dlx8 = diff(log(df[,9]))
dlx9 = diff(log(df[,10]))
ts.plot(dlx1,ylab="",xlab="",main="原油布蘭特北海現貨價(美元/桶)")
ts.plot(dlx2,ylab="",xlab="",main="布蘭特原油倫敦ICE近月期貨收盤價(美元/桶)")
ts.plot(dlx3,ylab="",xlab="",main="輕油(石油腦)日本C＆F(美元/公噸)")
ts.plot(dlx4,ylab="",xlab="",main="乙烯遠東區CFR(美元/公噸)")
ts.plot(dlx5,ylab="",xlab="",main="丙烯東南亞CFR(美元/公噸)")
ts.plot(dlx6,ylab="",xlab="",main="苯東南亞CFR(美元/公噸)")
ts.plot(dlx7,ylab="",xlab="",main="低密度聚乙烯LDPE通用級遠東區CFR(美元/公噸)")
ts.plot(dlx8,ylab="",xlab="",main="高密度聚乙烯HDPE射出級遠東區CFR(美元/公噸)")
ts.plot(dlx9,ylab="",xlab="",main="聚丙烯PP吹袋級遠東區CFR(美元/公噸)")


#檢查是否具有單位根
library(tseries)

adf.test(x1,k=0)
adf.test(x1)
pp.test(x1)

#有遺漏值無法做單根檢定，補充遺漏值
install.packages("imputeTS")
library(imputeTS)

filled_x1 <- na.interpolation(x1, option = "linear")
filled_x2 <- na.interpolation(x2, option = "linear")
filled_x3 <- na.interpolation(x3, option = "linear")
filled_x4 <- na.interpolation(x4, option = "linear")
filled_x5 <- na.interpolation(x5, option = "linear")
filled_x6 <- na.interpolation(x6, option = "linear")
filled_x7 <- na.interpolation(x7, option = "linear")
filled_x8 <- na.interpolation(x8, option = "linear")
filled_x9 <- na.interpolation(x9, option = "linear")

any(is.na(filled_x1))
any(is.na(filled_x2))
any(is.na(filled_x3))
any(is.na(filled_x4))
any(is.na(filled_x5))
any(is.na(filled_x6))
any(is.na(filled_x7))
any(is.na(filled_x8))
any(is.na(filled_x9))

#檢查補漏值後的走勢圖
ts.plot(filled_x1,ylab="",xlab="",main="原油布蘭特北海現貨價(美元/桶)")
ts.plot(filled_x2,ylab="",xlab="",main="布蘭特原油倫敦ICE近月期貨收盤價(美元/桶)")
ts.plot(filled_x3,ylab="",xlab="",main="輕油(石油腦)日本C＆F(美元/公噸)")
ts.plot(filled_x4,ylab="",xlab="",main="乙烯遠東區CFR(美元/公噸)")
ts.plot(filled_x5,ylab="",xlab="",main="丙烯東南亞CFR(美元/公噸)")
ts.plot(filled_x6,ylab="",xlab="",main="苯東南亞CFR(美元/公噸)")
ts.plot(filled_x7,ylab="",xlab="",main="低密度聚乙烯LDPE通用級遠東區CFR(美元/公噸)")
ts.plot(filled_x8,ylab="",xlab="",main="高密度聚乙烯HDPE射出級遠東區CFR(美元/公噸)")
ts.plot(filled_x9,ylab="",xlab="",main="聚丙烯PP吹袋級遠東區CFR(美元/公噸)")

filled_lx1 = log(filled_x1)
filled_lx2 = log(filled_x2)
filled_lx3 = log(filled_x3)
filled_lx4 = log(filled_x4)
filled_lx5 = log(filled_x5)
filled_lx6 = log(filled_x6)
filled_lx7 = log(filled_x7)
filled_lx8 = log(filled_x8)
filled_lx9 = log(filled_x9)
ts.plot(filled_lx1,ylab="",xlab="",main="原油布蘭特北海現貨價(美元/桶)")
ts.plot(filled_lx2,ylab="",xlab="",main="布蘭特原油倫敦ICE近月期貨收盤價(美元/桶)")
ts.plot(filled_lx3,ylab="",xlab="",main="輕油(石油腦)日本C＆F(美元/公噸)")
ts.plot(filled_lx4,ylab="",xlab="",main="乙烯遠東區CFR(美元/公噸)")
ts.plot(filled_lx5,ylab="",xlab="",main="丙烯東南亞CFR(美元/公噸)")
ts.plot(filled_lx6,ylab="",xlab="",main="苯東南亞CFR(美元/公噸)")
ts.plot(filled_lx7,ylab="",xlab="",main="低密度聚乙烯LDPE通用級遠東區CFR(美元/公噸)")
ts.plot(filled_lx8,ylab="",xlab="",main="高密度聚乙烯HDPE射出級遠東區CFR(美元/公噸)")
ts.plot(filled_lx9,ylab="",xlab="",main="聚丙烯PP吹袋級遠東區CFR(美元/公噸)")

filled_dlx1 = diff(log(filled_lx1))
filled_dlx2 = diff(log(filled_lx2))
filled_dlx3 = diff(log(filled_lx3))
filled_dlx4 = diff(log(filled_lx4))
filled_dlx5 = diff(log(filled_lx5))
filled_dlx6 = diff(log(filled_lx6))
filled_dlx7 = diff(log(filled_lx7))
filled_dlx8 = diff(log(filled_lx8))
filled_dlx9 = diff(log(filled_lx9))
ts.plot(filled_dlx1,ylab="",xlab="",main="原油布蘭特北海現貨價(美元/桶)")
ts.plot(filled_dlx2,ylab="",xlab="",main="布蘭特原油倫敦ICE近月期貨收盤價(美元/桶)")
ts.plot(filled_dlx3,ylab="",xlab="",main="輕油(石油腦)日本C＆F(美元/公噸)")
ts.plot(filled_dlx4,ylab="",xlab="",main="乙烯遠東區CFR(美元/公噸)")
ts.plot(filled_dlx5,ylab="",xlab="",main="丙烯東南亞CFR(美元/公噸)")
ts.plot(filled_dlx6,ylab="",xlab="",main="苯東南亞CFR(美元/公噸)")
ts.plot(filled_dlx7,ylab="",xlab="",main="低密度聚乙烯LDPE通用級遠東區CFR(美元/公噸)")
ts.plot(filled_dlx8,ylab="",xlab="",main="高密度聚乙烯HDPE射出級遠東區CFR(美元/公噸)")
ts.plot(filled_dlx9,ylab="",xlab="",main="聚丙烯PP吹袋級遠東區CFR(美元/公噸)")

#檢查單位根
adf.test(filled_x1)

adf.test(filled_lx1)

adf.test(filled_dlx1)$p.value

cbind(adf.test(filled_dlx1)$p.value,adf.test(filled_dlx2)$p.value,adf.test(filled_dlx3)$p.value,adf.test(filled_dlx4)$p.value,adf.test(filled_dlx5)$p.value,adf.test(filled_dlx6)$p.value,adf.test(filled_dlx7)$p.value,adf.test(filled_dlx8)$p.value,adf.test(filled_dlx9)$p.value)
##dlx全都顯著，拒絕h0，代表資料具平穩性

###做VAR模型###
install.packages("vars")
library("vars")

filled_dlx = cbind(filled_dlx1,filled_dlx2,filled_dlx3,filled_dlx4,filled_dlx5,filled_dlx6,filled_dlx7,filled_dlx8,filled_dlx9)

VARselect(filled_dlx, lag.max=10, type="both") #AIC選p=9,BIC選p=7，故選p=7。

VAR_model = VAR(filled_dlx,p = 7)
VAR_model

summary(VAR_model)

mq(VAR_model$residuals,adj = 9^2 * 7)#adj為扣除k^2 * p 之自由度部分，檢查發現p<7部分皆不顯著，代表模型還可以。
