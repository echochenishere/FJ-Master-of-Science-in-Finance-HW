require("pracma")
tic()
##資料載入
setwd("C:\\Users\\User\\Desktop")
return=read.csv("return.csv",sep=",",header=TRUE)
#head(return);tail(return);colnames(return)

#設定想要的投組個數(由亂數產生器產生，均等分配，(-4,4))
number=10
weights=matrix(runif((ncol(return)-2)*number,min = -4, max = 4), nrow=number)
weights=cbind(weights,1-rowSums(weights))
colnames(weights)=colnames(return)[2:ncol(return)]

#計算個股平均&共變異數矩陣
#平均
stock_mean=colMeans(return[,2:ncol(return)]);stock_mean=cbind(stock_mean)
##變異數
##colVars = function(x, na.rm=FALSE, dims=1, unbiased=TRUE, SumSquares=FALSE,twopass=FALSE) { 
##  if (SumSquares) return(colSums(x^2, na.rm, dims)) 
##  N <- colSums(!is.na(x), FALSE, dims) 
##  Nm1 <- if (unbiased) N-1 else N 
##  if (twopass) {x <- if (dims==length(dim(x))) x - mean(x, na.rm=na.rm) else 
##  sweep(x, (dims+1):length(dim(x)), colMeans(x,na.rm,dims))} 
##  (colSums(x^2, na.rm, dims) - colSums(x, na.rm, dims)^2/N) / Nm1 
##} 
#var=colVars(return[,2:ncol(return)])
#共變異數矩陣
stock_cov_matrix=cov(return[,2:ncol(return)])

#投資組合平均&變異數&標準差
#平均
p_mean=weights%*%stock_mean;colnames(p_mean)="p_mean"
p_var=c()
for(i in 1:number){
      p_var[i]=weights[i,]%*%stock_cov_matrix%*%t(weights)[,i]
      p_var=cbind(p_var);colnames(p_var)="p_var"
      }
p_std=p_var^0.5;colnames(p_std)="p_std"
portfolio=cbind(p_var,p_std,p_mean)

#效率前緣公式
one_vector=matrix(1,nrow=ncol(return)-1,ncol = 1)
A=t(one_vector)%*%solve(stock_cov_matrix)%*%stock_mean;colnames(A)="A";rownames(A)="formula"
B=t(stock_mean)%*%solve(stock_cov_matrix)%*%stock_mean;colnames(B)="B";rownames(B)="formula"
C=t(one_vector)%*%solve(stock_cov_matrix)%*%one_vector;colnames(C)="C";rownames(C)="formula"
D=B*C-A^2;colnames(D)="D";rownames(D)="formula"
g=(B[1,1]*(solve(stock_cov_matrix)%*%one_vector      )-A[1,1]*(solve(stock_cov_matrix)%*%stock_mean ))/D[1,1];colnames(g)="g"
h=(C[1,1]*(solve(stock_cov_matrix)%*%stock_mean      )-A[1,1]*(solve(stock_cov_matrix)%*%one_vector ))/D[1,1];colnames(h)="h"
#colSums(g);colSums(h)
E_mean=p_mean;colnames(E_mean)="E_mean"
E_var=(C[1,1]/D[1,1])*((E_mean-A[1,1]/C[1,1])^2)+(1/C[1,1]);colnames(E_var)="E_var"
E_std=E_var^0.5;colnames(E_std)="E_std"
efficient_frontier=cbind(E_var,E_std,E_mean)

#最小變異數投資組合mvp
mvp_mean=A[1,1]/C[1,1]
mvp_var=(1/C[1,1])
mvp_std=(1/C[1,1])^0.5
mvp=cbind(mvp_var,mvp_std,mvp_mean);rownames(mvp)="mvp"

#結果匯出
mean_and_cov_matrix=cbind(stock_mean,stock_cov_matrix)
formula1=cbind(A,B,C,D)
formula2=cbind(g,h)
result1=cbind(weights,portfolio,efficient_frontier)
result2=mvp
save.xlsx <- function (file, ...)
  {
      require(xlsx, quietly = TRUE)
      objects <- list(...)
      fargs <- as.list(match.call(expand.dots = TRUE))
      objnames <- as.character(fargs)[-c(1, 2)]
      nobjects <- length(objects)
      for (i in 1:nobjects) {
          if (i == 1)
              write.xlsx(objects[[i]], file, sheetName = objnames[i])
          else write.xlsx(objects[[i]], file, sheetName = objnames[i],
              append = TRUE)
      }
      print(paste("Workbook", file, "has", nobjects, "worksheets."))
}
save.xlsx("efficient_frontier.xlsx",mean_and_cov_matrix,formula1,formula2,result1,result2)

#畫圖
plot(E_std,E_mean,type="p",main="efficient frontier",xlab="std",ylab="mean",pch=16,col="red")
points(p_std  ,p_mean  ,type="p",pch=17,col="blue")
points(mvp_std,mvp_mean,type="p",pch=18,col="white")
dev.print(pdf, 'efficient frontier.pdf')
#散佈圖參考：http://www.math.nsysu.edu.tw/~lomn/homepage/R/R_plot.htm
toc()
