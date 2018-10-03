##################################################################################################
# NVT script for co-located + constrained + standard (i.e. all types of) trials
#
##################################################################################################
#
#                                **** SS Setup Script v3 ****
#
#Date: September 2017
#Author: DTolhurst with input from BCullis, KMathews, ASmith, CLisle
#
##################################################################################################

# **** standard + co-located trials (FIELD LAYOUT KNOWN or UNKNOWN)****
# terminology...
# co-located trials: more than one experiment at a location where 
# Same plot sizes, sowing times are sufficiently close and agronomic practices similar.

# COL: co-located trials with layouts KNOWN
# CON: co-located trials with layouts UNKNOWN and hence variance parameters 
# are to be 'constrained' equal 
# STD: standard trials (i.e. single experiment at a location, NOT co-located)
##################################################################################################

# This is the setup script used in KNVT

#SETUP
source("NVTfns5.6.R")
setup.fn(computer = 'linux') # enter computer if not a mac

data.temp <- data.read(infile="TrialExport.xls")
wheatset1.temp <- data.temp$data
head(wheatset1.temp)
(na.harvest <- data.temp$nas)

levels(factor(wheatset1.temp$Experiment))
nlevels(factor(wheatset1.temp$Experiment))
levels(factor(wheatset1.temp$Location))
nlevels(factor(wheatset1.temp$Location))
levels(factor(wheatset1.temp$Environment))
nlevels(factor(wheatset1.temp$Environment))
levels(factor(wheatset1.temp$Variety))
nlevels(factor(wheatset1.temp$Variety))
head(wheatset1.temp)

# only take complete constrained co-located trials
ExptTypes.df <- read.csv("../1Layouts/ConstrainedExpts.csv", strip.white = TRUE)
constrained <- ExptTypes.df[ExptTypes.df$e1 %in% wheatset1.temp$Experiment & ExptTypes.df$e2 %in% wheatset1.temp$Experiment,]
not.ready <- as.vector(as.matrix(ExptTypes.df[!ExptTypes.df$Environment %in% constrained$Environment,c("e1","e2"),]))
constrained <- as.vector(as.matrix(constrained[,c("e1","e2")]))
wheatset1.temp <- droplevels(wheatset1.temp[!wheatset1.temp$Experiment %in% not.ready,])

wheatset1.temp$YrCon <- as.character(wheatset1.temp$Experiment)

# co-located trials with layouts KNOWN, series is E and/or M, change to "M" when doing data for paper
layoutinfile="../1Layouts/LayoutData2017.csv"
unique(wheatset1.temp$Experiment)
temp <- layouts.read(infile=layoutinfile, data=wheatset1.temp)
layouts.temp <- temp$Layouts
colocated <- temp$Experiments
layoutdata.df <- temp$Data
unique(layoutdata.df$Experiment)

(standard <- levels(factor(layoutdata.df$Experiment[!layoutdata.df$Experiment %in% c(colocated, constrained)])))

#coding for expts that are standard while others at the same environment are co-located
tt <- layoutdata.df$Experiment %in% standard & layoutdata.df$Environment %in% unique(layoutdata.df$Environment[layoutdata.df$Experiment %in% constrained])
layoutdata.df$Environment <- as.character(layoutdata.df$Environment)
layoutdata.df$YrCon[tt] <- layoutdata.df$Environment[tt] <- paste(layoutdata.df$Environment[tt], layoutdata.df$Series[tt],sep="-")

#coding for expts that are co-located, but not constrained
tt2 <- with(layoutdata.df, table(Environment, Experiment))
tt2[tt2>1] <- 1
tt3 <- layoutdata.df$Environment %in%  rownames(tt2)[rowSums(tt2)>1]
tt4 <- tt3 & layoutdata.df$Experiment %in% standard 
layoutdata.df$YrCon[tt4] <- layoutdata.df$Environment[tt4] <- paste(layoutdata.df$Environment[tt4], layoutdata.df$Series[tt4],sep="-")
layoutdata.df$YrCon[layoutdata.df$Experiment %in% colocated] <- layoutdata.df$Environment[layoutdata.df$Experiment %in% colocated]

standard <- unique(c(standard, levels(factor(layoutdata.df$Experiment[layoutdata.df$Experiment %in% colocated & 
                                                                      layoutdata.df$Environment %in% rownames(tt2)[rowSums(tt2)==1]]))))
#coding for std expts
layoutdata.df$YrCon[layoutdata.df$Experiment %in% standard] <- layoutdata.df$Environment[layoutdata.df$Experiment %in% standard]

colocated <- colocated[!colocated %in% standard]
all.lays <- data.frame(Experiment = c(colocated, constrained, standard),
                       TrialType = rep(c("COL","CON","STD"), c(length(colocated),length(constrained),length(standard))))
if(length(layouts.temp)>1){layouts.temp <- layouts.temp[!layouts.temp$e1 %in% standard,]
                            rownames(layouts.temp) <- NULL}

# run the colocated layout function
unique(layoutdata.df$Experiment)
wheatset1.temp2 <- layouts.build(data=layoutdata.df, layouts=layouts.temp)
wheatset1.df <- wheatset1.temp2$data
(var.exclude <- wheatset1.temp2$var.exclude)

unique(wheatset1.df$Experiment)
head(wheatset1.df)
str(wheatset1.df)
t.sum <- trials.sum(wheatset1.df)

##################################################################################################
# Ready to Rock and Roll
##################################################################################################

# Check that it is ordered by year location
(ee <- levels(factor(wheatset1.df$Experiment)))
(ne <- length(ee))
(ll <- levels(factor(wheatset1.df$Location)))
(nl <- length(ll))
(yy <- levels(factor(wheatset1.df$Environment)))
(ny <- length(yy))
(yc <- levels(wheatset1.df$YrCon))
(nc <- length(yc))
(gg <- levels(factor(wheatset1.df$Variety)))
(ng <- length(gg))
levels(wheatset1.df$Range)
levels(wheatset1.df$Row)

##################################################################################################
# check reps are aligned with logical blocks...
##################################################################################################
# check that reps are in same direction for co-located trials.
rep.temp <- reps.check(data=wheatset1.df, BlockID="Rep", outfile1="1Graphics/wheatset1-layout-RepOrig.pdf",
                       outfile2="1Graphics/wheatset1-VarietyFreqperRepOrig.pdf")
(rep.na <- rep.temp$rep.na)
(rep.bad <- rep.temp$rep.bad)

##################################################################################################
# check rowreps are aligned with logical blocks...
##################################################################################################
# check that rowreps are in same direction for co-located trials.
rrep.temp <- reps.check(data=wheatset1.df, BlockID="RowRep", outfile1="1Graphics/wheatset1-layout-RowRepOrig.pdf",
                       outfile2="1Graphics/wheatset1-VarietyFreqperRowRepOrig.pdf")
(rrep.na <- rrep.temp$rep.na)
(rrep.bad <- rrep.temp$rep.bad)


allreps.na <- data.frame(Experiment=ee, Rep=0, RowRep=0)
allreps.na$Rep[allreps.na$Experiment %in% rep.na] <- 1
allreps.na$RowRep[allreps.na$Experiment %in% rrep.na] <- 1

allreps.bad <- data.frame(Experiment=ee, Rep=0, RowRep=0)
allreps.bad$Rep[allreps.bad$Experiment %in% rep.bad] <- 1
allreps.bad$RowRep[allreps.bad$Experiment %in% rrep.bad] <- 1

##################################################################################################
# Print out layouts
##################################################################################################
layouts.print(data=wheatset1.df, outfile="1Graphics/wheatset1-ColocatedLayouts.pdf")


##################################################################################################
# Check Out the Raw Data
##################################################################################################

# raw data plots
# print to pdf
pdf("1Graphics/wheatset1-Yield.pdf")
for(i in levels(wheatset1.df$Experiment)){
  print(ggplot(data=droplevels(subset(wheatset1.df,Experiment==i & !Variety %in% var.exclude)), aes(x=Variety, y=yield, label=Range)) + geom_text() + ggtitle(i) +
        stat_summary(fun.y = mean, fun.ymin = min, fun.ymax=max, colour='red') + coord_flip() +
        theme(axis.text.y = element_text(size=8, angle=30), title = element_text(size=9)))
}
dev.off()

##################################################################################################
#Generate the models.csv File
##################################################################################################

models <- cbind.data.frame(unique(wheatset1.df[, c("YrCon","Environment", "Experiment")])[,c("Environment", "Experiment")],
                           rep=rep(1,ne),
                           rrep=rep(1,ne),
                           lrow=rep(0,ne),
                           lcol=rep(0,ne),
                           rrow=rep(0,ne),
                           rcol=rep(0,ne),
                           resid="aa",
                           YrCon = unique(wheatset1.df[, c("YrCon","Environment", "Experiment")])[,"YrCon"])
rownames(models) <- NULL

#Now to customising for this dataset
#Now to specifying the residuals term in the models file.
# This needs to be altered according to Environment basis
models$resid <- as.character(models$resid)
(env.ia <- as.character(t.sum$Environment[t.sum$Nrange < 4 & t.sum$Nrow > 3])) #this assumes Range:Row. fit this for trials with only 3 ranges. None in this dataset.
(env.ai <- as.character(t.sum$Environment[t.sum$Nrange > 3 & t.sum$Nrow < 4]))
(env.ii <- as.character(t.sum$Environment[t.sum$Nrange < 4 & t.sum$Nrow < 4]))
models$resid[models$Environment%in%env.ia] <- "ia"
models$resid[models$Environment%in%env.ai] <- "ai"
models$resid[models$Environment%in%env.ii] <- "ii"
models$resid <- factor(models$resid)

##################################################################################################
# See what Covariates there are
##################################################################################################

str(wheatset1.df)
cov.temp <- covariates.print(data=wheatset1.df, infile="../1Functions/NVTCovariates2017.csv", 
                             outfile="1Graphics/wheatset1-Covariates.pdf")
(cov.df <- cov.temp$cov.df)
wheatset1.df <- cov.temp$data
str(wheatset1.df)
##################################################################################################
# write data out 1
##################################################################################################

# remove YrCon for those crops that dont need it...
if(length(grep("Q|B|W|L|P|O", substring(wheatset1.df$Experiment,1,1))) > 0 &
   length(c(colocated,constrained))<1){
  wheatset1.df <- wheatset1.df[,grep("YrCon",colnames(wheatset1.df),invert=TRUE)]
  models <- models[,grep("YrCon",colnames(models),invert=TRUE)]}

if(length(na.harvest)>0){
  write.csv((na.harvest), file="1DataFromKNVT/wheatset1-NAharvest.csv", row.names = FALSE)
}

write.csv(models, file="1DataFromKNVT/wheatset1-models.csv", row.names = FALSE)

write.csv(allreps.na, file="1DataFromKNVT/wheatset1-NAreps.csv", row.names = FALSE)

write.csv(allreps.bad, file="1DataFromKNVT/wheatset1-BADreps.csv", row.names = FALSE)

if(length(cov.df)>0){
  write.csv(cov.df, file="1DataFromKNVT/wheatset1-Covariates.csv", row.names = FALSE)
}

write.csv(wheatset1.df, file="1DataFromKNVT/wheatset1-DataFromKNVT.csv", row.names = FALSE)

write.csv(wheatset1.comments, file="1DataFromKNVT/wheatset1-Comments.csv", row.names = FALSE)

write.csv(all.lays, file="1DataFromKNVT/wheatset1-TrialTypes.csv", row.names = FALSE)

# end of setup script used in KNVT
#
#
# Start of setup script used outside KNVT (by Biometricians)
#
##################################################################################################
# fix reps + rowreps...
##################################################################################################
# check trial comments
# check out na reps/rowreps
allreps.na

# # check out bad reps/rowreps
# allreps.bad
# temp <- wheatset1.df
# temp$Rep[temp$Experiment == "DMKA17KUNR6"] <- 1
# temp$Rep[temp$Experiment == "DMKA17KUNR6" & temp$Range %in% 4:6] <- 2
# # temp$Rep[temp$Experiment == "DMKA17KUNR6" & temp$Range %in% 5:6] <- 3
# with(droplevels(subset(temp, Experiment=="DMKA17KUNR6")), tapply(Rep, list(Row,Range), unique))
# with(droplevels(subset(temp, Experiment=="DMKA17KUNR6")), table(Variety,Rep))
# # now to actual data
# wheatset1.df$Rep[wheatset1.df$Experiment == "OMaA17DAND6"] <- 1
# wheatset1.df$Rep[wheatset1.df$Experiment == "OMaA17DAND6" & wheatset1.df$Range %in% 3:4] <- 2
# wheatset1.df$Rep[wheatset1.df$Experiment == "OMaA17DAND6" & wheatset1.df$Range %in% 5:6] <- 3
# with(droplevels(subset(wheatset1.df, Experiment=="OMaA17DAND6")), tapply(Rep, list(Row,Range), unique))
# with(droplevels(subset(wheatset1.df, Experiment=="OMaA17DAND6")), table(Variety,Rep))

# which experiments not to include reps/rowreps because they are NA/cannot be fixed
rep.rm <- logical(0) # "OMaA17DAND6"
rrep.rm <- logical(0)

# Now, we NA those reps/rowreps that are NOT to be fitted.
# this ensures rep/rowrep is not fited in the future. 
# Also allows us to fit rep/rowrep to environment even when 
# some co-located experiments have NAs. 

allreps.rm <- data.frame(Experiment=ee, Rep=0, RowRep=0)
allreps.rm$Rep[allreps.rm$Experiment %in% rep.rm] <- 1
allreps.rm$RowRep[allreps.rm$Experiment %in% rrep.rm] <- 1

wheatset1.df$Rep[wheatset1.df$Experiment %in% allreps.rm$Experiment[allreps.rm$Rep==1]] <- NA
wheatset1.df$RowRep[wheatset1.df$Experiment %in% allreps.rm$Experiment[allreps.rm$RowRep==1]] <- NA

# fix in models.csv
models$rep[models$Experiment %in% allreps.rm$Experiment[allreps.rm$Rep==1]] <- 0
models$rrep[models$Experiment %in% allreps.rm$Experiment[allreps.rm$RowRep==1]] <- 0

save.image()


##################################################################################################
# See what Covariates I Fit
# Cross Reference with Trial Notes!!
##################################################################################################

# manually remove any variates
cov.df
# cov.df <- cov.df[!(cov.df$cov == "shatter"),] # variate not a covariate...
# cov.df <- cov.df[!(cov.df$cov == "nrows"),] # not a covariate...
# cov.df <- cov.df[!(cov.df$cov == "comments"),] # not a covariate...
# cov.df <- cov.df[!(cov.df$Experiment == "LMAA16GOON2" & cov.df$cov == "est"),]

# ok, now check whether these are genetically driven... # constrain variance paramters where apprpriate.
models
temp.asr2 <- asreml(est ~ Experiment + at(Experiment):Variety,
                   random= ~ at(Experiment):Rep + at(Experiment):RowRep, #change as required
                   residual = ~ ar1(Range):ar1(Row), #,#dsum(~id(Range):ar1(Row)| Environment),#change as required
                   na.action = na.method(x="include"),
                   data= droplevels(subset(wheatset1.df, Experiment=="OMaA17DAND6")))
temp.asr2 <- update(temp.asr2)
summary(temp.asr2)$varcom
wald(temp.asr2, denDF="default")
wheatset1.comments # fit miss at waik, est + earlygs at DAND 
cor(droplevels(subset(wheatset1.df, Experiment=="OMaA17DAND6"))$est,droplevels(subset(wheatset1.df, Experiment=="OMaA17DAND6"))$earlygs,use = "pairwise.complete.obs")

# If I choose to fit a covariate at only one Expt in an environment I will NA the other
# so we can still fit at(Environment):est

# genetically driven, remove est as covariate
test.cov.df <- cov.df[1:3,]
# test.cov.df <- logical(0) # remove est
  
##################################################################################################
# Customise the models.csv File
##################################################################################################

if(length(test.cov.df)>0){cov.mat <- matrix(0, nrow = nrow(models), ncol = length(unique(test.cov.df$cov)))
dimnames(cov.mat) <- list(NULL, unique(test.cov.df$cov))
models <- cbind.data.frame(models, cov.mat)}
rownames(models) <- NULL
head(models)

BLmodelsA <- models
# add in the ones we want
if(length(test.cov.df)>0){
for(i in 1: nrow(test.cov.df)){
  temp.cov <- droplevels(test.cov.df[i,])
  BLmodelsA[BLmodelsA$Experiment == as.character(temp.cov$Experiment),as.character(temp.cov$cov)] <- 1
}}
BLmodelsA

# make sure that only one residual term is fitted per Environment
with(unique(models[, c("Environment", "resid")]), table(Environment, resid))


##################################################################################################
# write data out 2
##################################################################################################

save(list=c("standard","constrained","colocated","trials.sum","var.exclude","allreps.rm","BLmodelsA","wheatset1.df"), file="1Data4SSAnalysis/wheatset1-Data4SSanalysis.RData")

save(list=ls(), file="1Data4SSAnalysis/wheatset1-allData4SSanalysis.RData")

write.csv(BLmodelsA, file="1Models/wheatset1-BLmodelsA.csv", row.names = FALSE)

write.csv(wheatset1.df, file="1Data4SSAnalysis/wheatset1-Data4SSanalysis.csv", row.names = FALSE)

write.csv(wheatset1.comments, file="1Data4SSAnalysis/wheatset1-Comments.csv", row.names = FALSE)

save.image()

#########################################################
#
#                     End of script
#
#                  Move to SSfit_DT.R
#
#########################################################
