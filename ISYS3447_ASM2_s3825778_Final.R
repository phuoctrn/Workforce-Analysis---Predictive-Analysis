library(car)
library(dplyr)
library(plyr)
library(fpp3)
library(tidyr)
library(pastecs)
library(VIM)
library(FactoMineR)
library(missMDA)
library(naniar)
library(strucchange)
library(corrplot)
library(GGally)
library(readxl)
library(janitor)
library(RColorBrewer)
library(ltm)
library(egg)
library(scales)
library(likert)
library(HH)
library(psych)
library(caret)
library(leaps)
library(fastDummies)
library(olsrr)
library(leaps)
library(tidyr)


#Import Data 

#Import data 
data_raw <- read_xlsx("/Users/phuoctran/Desktop/Intro to BA/ASM2/ISYS3447_A2_IntelliAuto-1.xlsx",sheet = 2) %>% subset(select = -IdNum)


###########Data Pre-processing#########

#Count missing value
is.na(data_raw) %>% count()

#Missing Value Visualization
vis_miss(data_raw) + coord_flip()

#Drop rows with NA values
data_raw <- data_raw %>% drop_na()

#Outlier detection using boxplots
data_numeric <- dplyr::select_if(data_raw, is.numeric)
par(mfrow=c(3,3))
boxplot(data_numeric$WorkHrs,horizontal = TRUE, main = "WorkHrs")
boxplot(data_numeric$Age,horizontal = TRUE, main = "Age")
boxplot(data_numeric$EducYrs,horizontal = TRUE, main = "EducYears")
boxplot(data_numeric$PreTaxInc,horizontal = TRUE, main = "PreTaxINc")
boxplot(data_numeric$PreTaxFamInc,horizontal = TRUE, main = "PreTaxFamInc")
boxplot(data_numeric$WrkYears,horizontal = TRUE, main = "WrkYears")
boxplot(data_numeric$EmpYears,horizontal = TRUE, main = "EmpYears")
boxplot(data_numeric$Engagement,horizontal = TRUE, main = "Engagement")

###########Question 1##############
## Advancement by genders
##Percentage stacked bar chart
##Percentage
advance_sex_prop <- prop.table(with(data_raw, table(Sex, Advances)), 1) %>% as.data.frame()
ad_sex_1 <- ggplot(advance_sex_prop,aes(x = Sex,y= Freq ,fill = Advances)) + 
  scale_fill_brewer(palette = "Blues",direction = -1)+
  geom_col(position = "stack",width = 0.6, color ="black")+
  geom_text(aes(label = percent(round(Freq,digits = 4))),position = position_stack(vjust = 0.5),color = "black",size =5,vjust = -3) +
  theme_classic() + scale_y_continuous(labels = scales::percent) +
  ylab("Percentage") + xlab("Genders") + ggtitle("") +
  guides(fill=guide_legend(title="Answers"))+
  theme(axis.text = element_text(size = 12),legend.position = "none") + coord_flip()

##Quantity stacked bar chart
advance_sex <- tabyl(data_raw, Sex, Advances) %>% pivot_longer(-1)
ad_sex_2 <- ggplot(advance_sex,aes(x=Sex,y=value,fill=name)) + 
  geom_col(position = "stack",width = 0.6,color = "black") + theme_classic() +
  scale_fill_brewer(palette = "Blues",direction = -1)+
  geom_text(aes(label=value),position = position_stack(vjust = 0.5),color = "black",size =5,check_overlap = TRUE,vjust = -3) +
  guides(fill=guide_legend(title="Answers")) + 
  ylab("Number of employees") + xlab("Genders") + 
  ggtitle("") + coord_flip() + ggtitle("(Percentage) stacked bar chart illustrate different answers between genders regarding career advancement")
  theme(axis.text = element_text(size = 12))

ggarrange(ad_sex_2, ad_sex_1, nrow = 2)

##Advancement by age groups
data_raw["age_group"] = cut(data_raw$Age, c(18, 29, 39, 49, 59, Inf), 
                            c("18-29", "30-39", "40-49", "50-59",">=60"), include.lowest=TRUE)

advance_age_prop <-prop.table(with(data_raw, table(age_group, Advances)), 1) %>% as.data.frame()

ad_age_1 <- ggplot(advance_age_prop,aes(x = age_group,y= Freq ,fill = Advances)) + 
  geom_col(position = "stack",width = 0.6, color ="black")+
  scale_fill_brewer(palette = "Blues",direction = -1)+
  geom_text(data=subset(advance_age_prop,Freq != 0),aes(label = percent(round(Freq,digits = 4))),
            position = position_stack(vjust = 0.5),color = "black",size =5,vjust = -1.75) +
  theme_classic() + scale_y_continuous(labels = scales::percent) +
  ylab("Percentage") + xlab("Age Groups") +
  guides(fill=guide_legend(title="Answers")) +
  theme(axis.text = element_text(size = 12),legend.position = "none") + coord_flip()

advance_age <- tabyl(data_raw, age_group, Advances) %>% pivot_longer(-1)
ad_age_2 <- ggplot(advance_age,aes(x=age_group,y=value,fill=name)) + 
  geom_col(position = "stack",width = 0.6,color = "black") + theme_classic() +
  scale_fill_brewer(palette = "Blues",direction = -1)+
  geom_text(data=subset(advance_age,value != 0),aes(label=value),
            position = position_stack(vjust = 0.5),color = "black",size =5,check_overlap = TRUE,vjust = -1.75) +
  guides(fill=guide_legend(title="Answers")) + 
  ylab("Number of employees") + xlab("Age Groups") + 
  ggtitle("(Percentage) stacked bar chart illustrate different answers between age groups regarding career advancement")+
  theme(axis.text = element_text(size = 12))+ coord_flip()

ggarrange(ad_age_2, ad_age_1, nrow = 2)

###########Question 2########

##Numeric predictor selection

#Scatterplot of EmpYears vs Numeric Variables
s1 <- ggplot(data = data_raw, aes(y= EmpYears,x=WorkHrs)) + geom_point(alpha =0.5) + geom_smooth(method = "lm",color = "black")+
  ggtitle("EmpYears vs WorkHrs")
s2 <- ggplot(data = data_raw, aes(y= EmpYears,x=Age)) + geom_point(alpha =0.5) + geom_smooth(method = "lm")+
  ggtitle("EmpYears vs Age")
s3 <- ggplot(data = data_raw, aes(y= EmpYears,x=EducYrs)) + geom_point(alpha =0.5) + geom_smooth(method = "lm",color = "black")+
  ggtitle("EmpYears vs EducYrs")
s4 <- ggplot(data = data_raw, aes(y= EmpYears,x=Earners)) + geom_point(alpha =0.5) + geom_smooth(method = "lm",color = "black")+
  ggtitle("EmpYears vs Earners")
s5 <- ggplot(data = data_raw, aes(y= EmpYears,x=PreTaxInc)) + geom_point(alpha =0.5) + geom_smooth(method = "lm",color="black")+
  ggtitle("EmpYears vs PreTaxInc")
s6 <- ggplot(data = data_raw, aes(y= EmpYears,x=PreTaxFamInc)) + geom_point(alpha =0.5) + geom_smooth(method = "lm",color = "black")+
  ggtitle("EmpYears vs PreTaxFamInc")
s7 <- ggplot(data = data_raw, aes(y= EmpYears,x=WrkYears)) + geom_point(alpha =0.5) + geom_smooth(method = "lm")+
  ggtitle("EmpYears vs WrkYears")
s8 <- ggplot(data = data_raw, aes(y= EmpYears,x=Engagement)) + geom_point(alpha =0.5) + geom_smooth(method = "lm")+
  ggtitle("EmpYears vs Engagement")

ggarrange(s1,s2,s3,s5,s6,s7,s8,
          ncol = 4, nrow = 2)

#Engagement Log Transform
log_engage <- ggplot(data = data_raw, aes(y= EmpYears,x=log(Engagement))) + 
  geom_point(alpha =0.5) + geom_smooth(method = "lm", formula= y ~ log(x))+
  ggtitle("EmpYears vs log(Engagement)")


data_raw$Engagement2 <- data_raw$Engagement^2

quad_engage <- ggplot(data = data_raw, aes(y= EmpYears,x= Engagement2)) + 
  geom_point(alpha =0.5) + geom_smooth(method = "lm", formula = y ~ x+I(x^2))+
  ggtitle("EmpYears vs Engagement^2") + xlab("Engagement^2")

log_engage <- ggplot(data = data_raw, aes(y= EmpYears,x= log(Engagement))) + 
  geom_point(alpha =0.5) + geom_smooth(method = "lm", formula = y ~ log(x))+
  ggtitle("EmpYears vs Log(Engagement)") + xlab("Log(Engagement)")

ggarrange(s8,log_engage,
          ncol = 2, nrow = 1)

noquad <- lm(data = data_raw, EmpYears ~ Engagement )
quad <- lm(data = data_raw, EmpYears ~ Engagement + I(Engagement^2))
decay <- lm(data = data_raw, EmpYears ~ log(Engagement))

summary(decay)

par(mfrow=c(1,2))
plot(noquad$fitted.values,noquad$residuals,xlab = "Fitted values", ylab = "Residuals",
     main = "Residuals vs Fitted value plot 
     of Linear Model EmpYears ~ Engagement")
lines(lowess(noquad$fitted.values,noquad$residuals), col="red")
plot(quad$fitted.values,quad$residuals, xlab = "Fitted values", ylab = "Residuals",
    main = "Residuals vs Fitted value plot 
    of Quadratic Model EmpYears ~ Engagement + Engagement^2")
lines(lowess(quad$fitted.values,quad$residuals), col="red")





#Correlation Matrix
corrplot(cor(data_numeric),method="color",  
         type="lower", order="hclust", 
         addCoef.col = "white",
         tl.col="black", tl.srt=45,
         diag=FALSE,col = COL1("Blues",10),col.lim = c(-1,1),addgrid.col = "black")

 #Potential numeric predictors
ggarrange(s2,s7,log_engage,
          ncol = 3, nrow = 1)









#Categorical predictor selection

#Bivariate Analysis with categorical variables

data_cate <- dplyr::select_if(data_raw, is.character)


p1 <- ggplot(data = data_raw, aes(y= EmpYears,x=Occupn)) + geom_boxplot() + theme_minimal()

p2 <- ggplot(data = data_raw, aes(y= EmpYears,x=Sex)) + geom_boxplot()

p3 <- ggplot(data = data_raw, aes(y= EmpYears,x=JobSat)) + geom_boxplot()

p4 <- ggplot(data = data_raw, aes(y= EmpYears,x=RichWork)) + geom_boxplot()

p5 <- ggplot(data = data_raw, aes(y= EmpYears,x=JobChar)) + geom_boxplot() 

p6 <- ggplot(data = data_raw, aes(y= EmpYears,x=GetAhead)) + geom_boxplot()


p7  <- ggplot(data = data_raw, aes(y= EmpYears,x=MemUnion)) + geom_boxplot()

p8<-ggplot(data = data_raw, aes(y= EmpYears,x=FutPromo)) + geom_boxplot() 

p9<- ggplot(data = data_raw, aes(y= EmpYears,x=SexPromo)) + geom_boxplot()

p10<-ggplot(data = data_raw, aes(y= EmpYears,x=Advances)) + geom_boxplot()

p11<- ggplot(data = data_raw, aes(y= EmpYears,x=IDecide)) + geom_boxplot()

p12<-ggplot(data = data_raw, aes(y= EmpYears,x=OrgMoney)) + geom_boxplot()

pt1 <- ggarrange(p1,p2,p3,p4,p5,p6,p7,p8,p9,p10,p11,p12,
                 ncol = 3, nrow = 7)

p13 <-ggplot(data = data_raw, aes(y= EmpYears,x=ProudOrg)) + geom_boxplot()


p14 <-ggplot(data = data_raw, aes(y= EmpYears,x=StayOrg)) + geom_boxplot()

p15 <-ggplot(data = data_raw, aes(y= EmpYears,x=UnManRel)) + geom_boxplot()

p16 <-ggplot(data = data_raw, aes(y= EmpYears,x=CoWrkRel)) + geom_boxplot()


p17 <-ggplot(data = data_raw, aes(y= EmpYears,x=Schooling)) + geom_boxplot()

p18 <-ggplot(data = data_raw, aes(y= EmpYears,x=Training)) + geom_boxplot()

p19 <-ggplot(data = data_raw, aes(y= EmpYears,x=AwareI4.0)) + geom_boxplot()

p20 <-ggplot(data = data_raw, aes(y= EmpYears,x = NumPromo,group = NumPromo)) + geom_boxplot()+  ggtitle("EmpYears vs NumPromo Grouped Boxplots")+
  scale_x_continuous(breaks=seq(0,7,1))

p21 <-ggplot(data = data_raw, aes(y= EmpYears,x = Trauma,group = Trauma)) + geom_boxplot() +
  scale_x_continuous(breaks=seq(0,5,1))
p22 <-ggplot(data = data_raw, aes(y= EmpYears,x = Earners,group = Earners)) + geom_boxplot()

pt2 <- ggarrange(p13,p14,p15,p16,p17,p18,p19,p20,p21,p22,
                 ncol = 3, nrow = 4)



#Fitting regression model
#make categorical variables factors
data_fit <- data_raw
data_fit <- data_fit %>% mutate_if(is.character,as.factor)
data_fit$NumPromo <- data_fit$NumPromo %>% as.factor()
data_fit$Trauma <- data_fit$Trauma %>% as.factor()
data_fit$Earners <- data_fit$Earners %>% as.factor()


#Multiple Regression

fit <- lm(data = data_fit, EmpYears ~ WrkYears + Engagement + NumPromo )
summary(fit)

write.csv(tidy(summary(fit)),"/Users/phuoctran/Desktop/Intro to BA/ASM2/fit_summary.csv")

plot(fit)


#Prediction 
above_age55 <- subset(data_raw,data_raw$Age >= 55)

summary(above_age55$WrkYears) #median = 38

summary(above_age55$Engagement) #Median = 2.78

summary(above_age55$NumPromo) #Median = 2


newdata <- data.frame(NumPromo=factor("2", levels=c(0:7)),
                      WrkYears = 38,
                      Engagement = 2.78)
predict(fit,newdata = newdata) #18.966 ~ 19




above_empyear20 <- subset(data_raw,data_raw$EmpYears >= 19)
nrow(above_empyear20)

stat.desc(above_age55)



















