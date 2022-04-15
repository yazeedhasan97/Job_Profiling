rm(list = ls())
graphics.off()

# 1| Preparation ----------------------------------------------------------
# 1.1| Libraries ----------------------------------------------------------
myPackages   = c('haven', 'mclust', 'poLCA', 'nlsem', 'depmixS4', 'survival', 
                 'rpart.plot', 'survMisc', 'ggplot2', 'ggparty','partykit', 
                 'treeClust', 'flexmix', 'openxlsx', 'fitdistrplus', 
                 'tidyverse', 'reshape2', 'factoextra', 'ggpubr', 'tidyquant',
                 'DescTools')
notInstalled = myPackages[!(myPackages %in% rownames(installed.packages()))]
if(length(notInstalled)) {
  install.packages(notInstalled)
}

library(haven)        # Reads .dta files.
library(openxlsx)     # Exports the results to excel.
library(partykit)     # Tree-structured regression and classification models.
library(treeClust)    # Produces dissimilarities for tree-based clustering.
library(ggparty)      # Extends ggplot2 functionality to the partykit package. Customizable visualizations for trees. 
library(factoextra)   # Dendogram.
library(rpart.plot)   # Plots rpart trees.
library(flexmix)      # Fits discrete mixtures of regression models.
library(fitdistrplus) # Fits of a parametric distribution to non-censored or censored data.
library(poLCA)        # Estimation of latent class and latent class regression models for polytomous outcome variables.
library(nlsem)        # Estimation of Structural Equation Mixture Modeling (SEMM).
library(depmixS4)     # Estimation of dependent mixture models.
library(mclust)       # Model-based clustering, classification, and density estimation (based on finite normal mixture modelling).
library(survival)     # Survival analysis.
library(survMisc)     # Analysis of right censored survival data. Extends the methods available in 'survival'.
# library(LTRCtrees)    # Fits survivals trees for left-truncated and right censored (LTRC) data. 
# For right censored data, the tree algorithms in 'rpart' and 'partykit' can be used to fit a survival tree.
library(tidyverse)    # Data manipulation and graph generation.
library(reshape2)     # Data manipulation for graphs in ggplot and export data.
library(extrafont)    # Font according to Word.
library(tidyquant)    # For statistical moments.
library(DescTools)    # To calculate entropy.
library(ggpubr)       # Nice trees.

loadfonts(device = 'win') # Load fonts from windows.
options(scipen = 999) # Disable scientific notation.

# 1.2| Working directory --------------------------------------------------
setwd('C:/Users/User/Desktop/Mini Jordan/UNEP/Unemployment-Profile')
library(here)
#setwd(enc2native(here()))

# 1.3| Functions ----------------------------------------------------------
grapherFunction = function(gend, method, lastObs, clusterNum, wb) {
  newSheet   = addWorksheet(wb, sheetName = paste0(method, ', ', gend, ', ', lastObs), gridLines = F)
  # To export:
  # - Excel:
  rowLine    = 1
  jumpOfRows = 20
  jumpOfCols = 20
  colPlots   = 1
  colTables  = 10
  # - By plot:
  base_exp     = 1
  heightExp    = 1
  widthExp     = 1
  scale_factor = base_exp/widthExp
  
  # Upload data -----------------------------------------------------------
  data = read_dta('Input/Unemployment Data.dta')
  
  # Some corrections:
  data$NationalID_Number = as.character(data$NationalID_Number)
  
  # Step 1 ----------------------------------------------------------------
  # Identify those cases where the position T-1 was unemployed.
  data$RegisterAfterFired = 0
  data$LastJobC           = 1
  maxDate                 = max(data$end_date, na.rm = T)
  
  data = data %>% 
    mutate(RegisterAfterFired = ifelse(NationalID_Number == lead(NationalID_Number) & reason_suspension_tr %in% c('Resignation', 'Laid off'), 1, 0)) %>%
    mutate(RegisterAfterFired = lag(RegisterAfterFired)) %>%
    mutate_at(vars(RegisterAfterFired), ~replace(., is.na(.), 0)) 
  
  if(method == 'Method 3'){ # NEEDS MORE ATTINTION
    data = data %>% mutate(LastJobC = ifelse(last_ind == 1 & reason_suspension_tr %in% c('Resignation', 'Laid off'), 0, LastJobC))
    
    data$LastJobC[data$`_merge_with_mol`==2]         = 0
    data$experience[data$`_merge_with_mol`==2]       = 0
    data$end_date[data$`_merge_with_mol`==2]         = data$RegistrationdateintoNEES[data$`_merge_with_mol`==2]
    data$econ_activity_tr[data$`_merge_with_mol`==2] = ''
    
    dataTemp                    = data[data$LastJobC%in%c(0),]
    data$LastJobC               = 1
    dataTemp$unempl_spell       = maxDate-dataTemp$end_date
    dataTemp$RegisterAfterFired = 1
    
    data                        = rbind(data,dataTemp) # NEEDS MORE ATTINTION # not sure this is a correct step as it wil creaet duplicate data
  }
  
  # Step 2 ----------------------------------------------------------------
  # Add first valid employments.
  # data$validCase = data$RegisterAfterFired
  data           = data %>% mutate(validCase = ifelse(first_ind == 1, 1, RegisterAfterFired))
  dataS          = data %>% filter(validCase == 1 & !(econ_activity_tr %in% c('Public administration, defense, and social security')))
  
  # Keep only complete registers.
  dataS = dataS %>% filter(!is.na(unempl_spell) & unempl_spell > 0 & unempl_spell < 1800)
  
  # Take only one if needed.
  if(!lastObs == 'All'){ # NEEDS MORE ATTINTION # not moved to python since attached to the graph building
    dataS = dataS %>%
      group_by(dataS$NationalID_Number) %>% # this is for people with multible jobs # it explains the duplicate id in the data
      filter(row_number() == n())
  }
  
  # Step 3 ----------------------------------------------------------------
  # Creation of variables:
  # - Dummy for first job:
  dataS          = dataS %>% mutate(FirstJob = ifelse(first_ind == 1, 1, 0))
  dataS$FirstJob = factor(dataS$FirstJob)
  
  # - Experience:
  dataS$Experience = round(dataS$experience/365, 0)
  dataS = dataS %>% mutate(Experience = case_when(Experience == 0                    ~  0,
                                                  Experience == 1                    ~  1,
                                                  Experience >  1 & Experience <=  5 ~  5,
                                                  Experience >  5 & Experience <= 10 ~ 10,
                                                  Experience > 10 & Experience <= 15 ~ 15,
                                                  Experience > 15                    ~ 20))
  dataS$Experience = factor(dataS$Experience)
  
  # - Age:
  dataS$Age = factor(10 * round((dataS$job_search_start - dataS$JobSeekers_DateOfBirth)/3650, 0)) # Measured in decades.
  
  # - Governorate:
  dataS$Region = factor(dataS$Governorate)
  
  # - Disability:
  dataS$Disability = factor(dataS$Disabled_tr)
  
  # - Gender:
  dataS$Sex = factor(dataS$Name_tr)
  
  # - Education Level:
  dataS = dataS %>% mutate(EducationalAttainment = case_when(EducationalAttainment %in% c("High Diploma", "Bachelor", "Masters", "PhD") ~  "Bachelor or above",
                                                             EducationalAttainment %in% c("Vocational Training") ~  "Vocational Training",
                                                             EducationalAttainment %in% c("Middle Diploma") ~  "Middle Diploma",
                                                             EducationalAttainment %in% c("Secondary or Below")~ "Secondary or Below" ))
  
  dataS$Education = factor(dataS$EducationalAttainment)
  
  # - Year of unemployment:
  dataS$UnemploymentYear= 0 #as.numeric(format(dataS$job_search_start, format='%Y'))
  
  # - Same Job:
  dataS$SameJob = factor(dataS$rep_job, 
                         levels = c(0, 1), 
                         labels = c('No', 'Yes'))
  # dataS$SameJob= 0
  
  # - Wages:
  temp              = as.numeric(dataS$wage_adj)
  temp[is.na(temp)] = -1000
  t                 = kmeans(temp, 4)
  te                = aggregate(temp, by = list(t$cluster), FUN = mean)
  te                = te[order(te[,2]),]
  tempo             = cbind(te[,1],1:(dim(te)[1]))
  tempo             = tempo[order(tempo[,1]),]
  
  WageGroup = t$cluster
  WageGroup[t$cluster%in%c(1)] = tempo[1,2]-1
  WageGroup[t$cluster%in%c(2)] = tempo[2,2]-1
  WageGroup[t$cluster%in%c(3)] = tempo[3,2]-1
  WageGroup[t$cluster%in%c(4)] = tempo[4,2]-1
  
  dataS$WageGroup = factor(WageGroup, 
                           levels = 0:3, 
                           labels = c('No information', 'Low', 'Medium', 'High'))
  dataS$WageGroup=0
  
  # - Poverty:
  dataS$Poverty = dataS$poverty_score
  dataS$Poverty = 0
  
  # - Industry:
  dataS$Industry=dataS$econ_activity_tr
  if(method == 'Method 2'){
    dataS$Industry[dataS$Industry%in%c("")]=NA
  }
  
  dataS$Industry = factor(dataS$Industry)
  # dataS$Industry = ''
  
  # - Unemployment Spell:
  dataS$UnemploymentSpell = dataS$unempl_spell/30 # Measured in months.
  
  # Select right column:
  if(lastObs == 'All') { # NEEDS MORE ATTINTION # what is lastObs and what is its main Job
    dataS           = dataS[, c(63,65:77)]
    colnames(dataS) = c('LastJobC' ,'FirstJob', 'Experience', 'Age', 'Governorate','Disability', 'Gender', 'EducationLevel','UnemploymentYear','SameJob', 'WageGroup','Poverty', 'Industry','UnemploymentSpell')
  } else {
    dataS           = dataS[, c(63,65:78)]
    colnames(dataS) = c('LastJobC' ,'NationalID_Number','FirstJob', 'Experience', 'Age', 'Governorate','Disability', 'Gender', 'EducationLevel','UnemploymentYear','SameJob', 'WageGroup','Poverty', 'Industry','UnemploymentSpell')
  }
  
  # Step 4 ----------------------------------------------------------------
  # Gender tree:
  gen = unique(dataS$Gender)[if(gend=='Male'){1}else{2}]
  finalDataF=dataS[dataS$Gender%in%gen&dataS$SameJob%in%c('No'),]
  finalDataF$one=1
  finalDataF=finalDataF[complete.cases(finalDataF),]  # NEEDS MORE ATTINTION # the number of records at this state reduced to 1.5K from 64K due to empty povirty score
  controlBasic = rpart.control(minbucket = if(lastObs == 'One'){500}else{500}, cp = 0.0002, maxsurrogate = 0, usesurrogate = 0, xval = 10)
  controlBasic_ot = rpart.control(minbucket = if(lastObs == 'One'){50}else{50}, cp = 0.0002, maxsurrogate = 0, usesurrogate = 0, xval = 10)
  
  #First Job is being removed
  if(method == 'Method 1') {
    tr_UnemploymentSpell = rpart(formula = UnemploymentSpell ~Experience+Age+Governorate+Disability+Gender+EducationLevel+Industry,
                                 data = finalDataF, control = controlBasic)
  } else {
    if(method == 'Method 2') {
      tr_UnemploymentSpell = rpart(formula = Surv(UnemploymentSpell, LastJobC) ~ Experience+Age+Governorate+Disability+Gender+EducationLevel+Industry,
                                   data = finalDataF, control = controlBasic)
    } else {
      tr_UnemploymentSpell = rpart(formula = Surv(UnemploymentSpell, LastJobC) ~ Experience+Age+Governorate+Disability+Gender+EducationLevel,
                                   data = finalDataF, control = controlBasic)
    }
  } 
  
  
  # NEEDS MORE ATTINTION # we don't need the below part as we are using place holders (moreover, it won't bedetriminstic) # YOU NEED TO FOLLOW THE MOST FREQUANT DESTRIBUTION
  # Training material: # NEEDS MORE ATTINTION # No need for this as we are now using place holders (this is called imputaion and may not be as much acurate as need)
  tr_FirstJob = rpart(formula = FirstJob ~ Experience+Age+Governorate+Disability+Gender+EducationLevel+Industry,
                      data = finalDataF, control = controlBasic_ot)
  
  tr_Experience = rpart(formula = Experience ~ Age+Governorate+Disability+Gender+EducationLevel+Industry,
                        data = finalDataF, control = controlBasic_ot)
  
  tr_Age = rpart(formula = Age ~ Experience+Governorate+Disability+Gender+EducationLevel+Industry,
                 data = finalDataF, control = controlBasic_ot)
  
  tr_Governorate = rpart(formula = Governorate ~ Experience+Age+Disability+Gender+EducationLevel+Industry,
                         data = finalDataF, control = controlBasic_ot)
  
  tr_Disability = rpart(formula = Disability ~ Experience+Age+Governorate+Gender+EducationLevel+Industry,
                        data = finalDataF, control = controlBasic_ot)
  
  tr_EducationLevel = rpart(formula = EducationLevel ~ Experience+Age+Governorate+Disability+Gender+Industry,
                            data = finalDataF, control = controlBasic_ot)
  
  tr_Industry = rpart(formula = Industry ~ Experience+Age+Governorate+Disability+Gender+EducationLevel,
                      data = finalDataF, control = controlBasic_ot)
  
  # Plot tree -------------------------------------------------------------
  splitFun <- function(x, labs, digits, varlen, faclen){
    # Line jump:
    labs <- gsub(" = ", ":\n", labs)
    labs <- gsub(",", " ", labs)
    # Creation of labels:
    for(i in 1:length(labs)) {
      labs[i] <- paste(strwrap(labs[i], width = 15), collapse = "\n")
    }
    return(labs)
  }
  
  png(paste0('g1_',gen,'_',method,'_',lastObs,'.png'), width = 1920*1.5, height = 1080*1.5)
  rpart.plot(tr_UnemploymentSpell, split.fun = splitFun, 
             digits = 2, ni = T, extra = 100, # Extra = 101 for number of obs, ni = T to know the leaf.
             branch.lty = 3, shadow.col = 'gray', fallen.leaves = F) # nn = TRUE to number of each node.
  dev.off()
  
  fileName = paste0('g1_',gen,'_',method,'_',lastObs,'.png')
  insertImage(wb, 
              file = fileName, sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots,
              width = 1920, height = 1080, units = 'px')
  rowLine = rowLine + jumpOfRows
  
  # Dendrogram ------------------------------------------------------------
  rp = tr_UnemploymentSpell
  newdata = finalDataF
  type = 'where'
  rp$frame$yval <- 1:nrow(rp$frame)
  should.be.leaves <- which(rp$frame[, 1] == "<leaf>")
  
  leaves <- predict(rp, newdata = newdata, type = "vector")
  should.be.leaves <- which (rp$frame[,1] == "<leaf>")
  bad.leaves <- leaves[!is.element (leaves, should.be.leaves)]
  if (length (bad.leaves) > 0){
    u.bad.leaves <- unique (bad.leaves)
    u.bad.nodes <- row.names (rp$frame)[u.bad.leaves]
    all.nodes <- row.names (rp$frame)[rp$frame[,1] == "<leaf>"]
    is.descendant <- function (all.leaves, node) {
      if (length (all.leaves) == 0) return (logical (0))
      all.leaves <- as.numeric (all.leaves); node <- as.numeric (node)
      if (missing (node)) return (NA)
      result <- rep (FALSE, length (all.leaves))
      for (i in 1:length (all.leaves)) {
        result=node<all.leaves
      }
      return (result)
    }
    where.tbl <- table (rp$where)
    names (where.tbl) <- row.names (rp$frame)[as.numeric (names (where.tbl))]
    for (u in 1:length (u.bad.nodes)) {
      desc.vec <- is.descendant (all.nodes, u.bad.nodes[u])
      me <- where.tbl[all.nodes][desc.vec]
      winner <- names (me)[me == max(me)][1]
      leaves[leaves == u.bad.leaves[u]] <- which (row.names (rp$frame) == winner)
    }
    leaves[leaves%in%u.bad.leaves] <- NA
  }
  
  s=leaves
  
  #s=rpart.predict.leaves(tr_UnemploymentSpell, finalDataF, type = 'where')
  keyNodes=unique(s)
  keyNodes=keyNodes[!is.na(keyNodes)]
  finalDataF$node=s
  
  
  #########Block from here
  statMoments = finalDataF %>% group_by(node) %>% summarise(mean     = mean(UnemploymentSpell),
                                                            variance = sd(UnemploymentSpell),
                                                            skew     = skewness(UnemploymentSpell),
                                                            kurtosis = kurtosis(UnemploymentSpell),
                                                            p25      = quantile(UnemploymentSpell, 0.25),
                                                            p50      = quantile(UnemploymentSpell, 0.50),
                                                            p75      = quantile(UnemploymentSpell, 0.75),
                                                            entropy  = Entropy(UnemploymentSpell))
  statMoments = statMoments %>% mutate(mean     = scale(mean),
                                       variance = scale(variance),
                                       skew     = scale(skew),
                                       kurtosis = scale(kurtosis),
                                       p25      = scale(p25),
                                       p50      = scale(p50),
                                       p75      = scale(p75),
                                       entropy  = scale(entropy)) %>% as.data.frame()
  
  rownames(statMoments) = statMoments[,1]
  statMoments = statMoments %>% select(-node)
  d = statMoments %>% dist()
  ######################Here
  #Densities:
  # denMat = matrix(0,100,0)
  # 
  # for(i in keyNodes){
  #   dx     = density(finalDataF$UnemploymentSpell[finalDataF$node%in%c(i)])
  #   xnew   = seq(3,10,length.out = 100)
  #   t      = approx(dx$x,dx$y,xout=xnew)
  #   denMat = cbind(denMat,as.matrix(approx(dx$x,dx$y,xout=xnew)$y))
  # }
  # denMat[is.na(denMat)] = 10^-4
  # 
  # d = KLdiv(denMat, method = c('continuous'), eps = 10^-4, overlap = TRUE)
  # 
  # # Hierarchical clustering using Ward:
  # colnames(d) = keyNodes
  # rownames(d) = keyNodes
  # 
  # d   = as.dist(d)
  # 
  ################Here
  
  hc1 = hclust(d, method = 'ward.D') ####### Just a referance
  k   = min(length(keyNodes), clusterNum) # Number of clusters.
  
  # Plot: 
  p = fviz_dend(hc1, k = k, color_labels_by_k = TRUE, palette = 'Set1',
                rect = TRUE, lower_rect = -0.02, rect_border = 'Set1', rect_fill = TRUE) +
    labs(title    = 'Clustering of nodes',
         subtitle = 'Ward Clustering',
         x        = 'Decision tree nodes',
         y        = 'Distance') +
    theme_classic() +
    theme(legend.position = 'bottom',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_blank(),
          axis.line.x     = element_blank(),
          axis.line.y     = element_blank(),
          axis.ticks.x    = element_blank(),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill  = guide_legend(''),
           color = guide_legend(override.aes = list(size = scale_factor * 1)))
  
  fileName = paste0('g2_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots,
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  rowLine = rowLine + jumpOfRows
  
  sub_grp <- cutree(hc1, k = k)
  
  # Same categories:
  keyNodes = sort(keyNodes)
  sCat=cbind(sub_grp, keyNodes)
  rownames(sCat)=NULL
  colnames(sCat)=c('Cluster','node')
  
  # * Nice plot -------------------------------------------------------------
  finalDataF=merge(finalDataF, sCat) # Cluster here isn't the right one yet...
  temp=aggregate(finalDataF$UnemploymentSpell, by=list(finalDataF$Cluster), FUN=mean) # Here they're organized.
  temp=temp[order(temp[,2]),]
  tempo=cbind(temp[,1],1:(dim(temp)[1]))
  colnames(tempo)=c('Cluster','Clusters')
  sCat = merge(sCat, tempo)     # Fixing...
  sCat = sCat[,c(3,2)]
  colnames(sCat)[1] = 'Cluster' # Fixed!
  finalDataF=merge(finalDataF, tempo)
  finalDataF$Cluster=finalDataF$Clusters
  
  finalDataF$NameCluster=factor(finalDataF$Cluster, levels = 1:k, labels=paste('Cluster', 1:k))
  finalDataF$Unity='Cluster'
  finalDataF$one=1
  
  # Tree:
  niceTree <- as.party(tr_UnemploymentSpell)
  g <- ggparty(niceTree)
  aux = unique(g$data$id[g$data$kids == 0]) # Tree leaves.
  
  g <- g + geom_edge(size = 1, colour = 'grey')
  g <- g + geom_edge_label(colour = 'black', size = 2)
  # g <- g + geom_node_plot(gglist = list(geom_density(aes(x = if(method == 'Method 1') {get('UnemploymentSpell')} else {get('Surv(UnemploymentSpell, LastJobC).time')})), # UnemploymentSpell or Surv,
  #                                       scale_y_continuous(expand = expansion(mult = c(0, .05))),
  #                                       scale_x_continuous(expand = expansion(mult = c(0, .05))),
  #                                       theme_classic(),
  #                                       labs(x = 'Unemployment spell',
  #                                            y = 'Density'),
  #                                       theme(legend.position = 'bottom',
  #                                             text            = element_text(family = 'Georgia'),
  #                                             axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
  #                                             axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
  #                                             axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
  #                                             legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
  #                                             strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
  #                                             plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
  #                                             plot.subtitle   = element_text(size = 10, family = 'Georgia'),
  #                                             legend.key.size = unit(0.5 * scale_factor, 'cm'))),
  #                         scales             = 'fixed',    # y and x axis fixed for better comparability.
  #                         ids                = 'terminal', # Labels for leaves.
  #                         shared_axis_labels = TRUE,
  #                         shared_legend      = TRUE,
  #                         legend_separator   = FALSE)
  g <- g + geom_node_label(aes(col = splitvar), # Labels for parent nodes.
                           show.legend = F,
                           line_list   = list(aes(label = splitvar)),
                           line_gpar   = list(list(size = 10)),
                           ids         = 'inner') + # Labels for parental nodes.
    geom_node_label(aes(label = paste0('Spell Group 1 \n N = ', nodesize, '\nNode ', id)), # Labels for leaves from cluster 1.
                    fontface = 'bold',
                    ids      = sCat[sCat[,1]==1,2],
                    size     = 3,
                    nudge_y  = 0.02,
                    color    = '#E41A1C') + # Red.
    geom_node_label(aes(label = paste0('Spell Group 2 \n N = ', nodesize, '\nNode ', id)), # Labels for leaves from cluster 2.
                    fontface = 'bold',
                    ids      = sCat[sCat[,1]==2,2],
                    size     = 3,
                    nudge_y  = 0.02,
                    color    = '#4DAF4A') + # Green.
    geom_node_label(aes(label = paste0('Spell Group 3 \n N = ', nodesize, '\nNode ', id)), # Labels for leaves from cluster 3.
                    fontface = 'bold',
                    ids      = sCat[sCat[,1]==3,2],
                    size     = 3,
                    nudge_y  = 0.02,
                    color    = '#377EB8')# Blue.)
  
  
  for(i in 1:length(g[['data']][['breaks_label']])) {
    if(class(g[['data']][['breaks_label']][[i]]) == 'character' & length(g[['data']][['breaks_label']][[i]]) != 1) {
      g[['data']][['breaks_label']][[i]][1:(length(g[['data']][['breaks_label']][[i]])-1)] = paste0(g[['data']][['breaks_label']][[i]][1:(length(g[['data']][['breaks_label']][[i]])-1)], '\n')
    }
  }
  
  # Density by node:
  densityByNode = finalDataF %>% # filter(Cluster == 3) %>% mutate(node = factor(node)) %>%
    ggplot(aes(x = UnemploymentSpell, colour = NameCluster)) +
    facet_wrap(node~., ncol = length(aux)) +
    geom_density(show.legend = F) +
    labs(x     = 'Unemployment Spell') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05))) +
    scale_x_continuous(expand = expansion(mult = c(0, .05))) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm'),
          plot.margin = unit(c(0,0,0,0), 'lines'),
          strip.text.x = element_blank(),
          strip.background = element_blank()) +
    guides(colour = guide_legend('Cluster', override.aes = list(size = scale_factor * 1)))
  
  fileName = paste0('tree_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, plot = ggarrange(g, densityByNode, nrow = 2, heights = c(0.8, 0.2), align = "v"),
         width = 6 * 6, height = 4 * 4)
  insertImage(wb,
              file = fileName, sheet = paste0(method,', ', gend, ', ', lastObs),
              startRow = rowLine, startCol = colPlots,
              width = 6 * 6, height = 4 * 4)
  rowLine = rowLine + jumpOfRows*4
  
  # * Information about the tree ------------------------------------------
  # Once we've got which cluster corresponds to each node, we can create the
  # table for the first figure, this is:
  dataFromPlot = rpart.plot(tr_UnemploymentSpell, split.fun = splitFun, 
                            digits = 5, ni = T, extra = 100, # Extra = 101 for number of obs, ni = T to know the leaf.
                            branch.lty = 3, shadow.col = 'gray', fallen.leaves = F) # nn = TRUE to number of each node.
  tempModification = data.frame(change = dataFromPlot[['labs']]) %>% 
    separate(change, into = c('Prediction', 'Percentage'), sep = '\\s') %>% 
    add_column(n = row.names(as.data.frame(dataFromPlot[['labs']])), .before = 1)
  tempModification$Prediction = as.numeric(tempModification$Prediction) %>% 
    round(digits = 3)
  
  anotherTempModification = as.data.frame(rpart.rules(tr_UnemploymentSpell, cover = TRUE, nn = TRUE, digits = 4))
  if(method == 'Method 1') {
    anotherTempModification$UnemploymentSpell = as.numeric(anotherTempModification$UnemploymentSpell) %>% round(digits = 3)
  } else {
    anotherTempModification$Surv = as.numeric(anotherTempModification$Surv) %>% round(digits = 3)
  }
  
  dataTree = merge(x = tempModification, y = anotherTempModification, by.x = 'Prediction', by.y = if(method == 'Method 1') {'UnemploymentSpell'}else{'Surv'}) %>% 
    merge(y = sCat, by.x = 'n', by.y = 'node') %>% dplyr::select(n, Cluster, `  cover`, everything(), -Percentage, -nn, -Prediction)
  colnames(dataTree)[4:length(dataTree)] = ' '
  
  writeData(wb, 
            sheet    = paste0(method,', ', gend, ', ', lastObs), 
            x        = dataTree, 
            startRow = 1, startCol = colTables)
  
  # Agreggate of cluster:
  p = ggplot(finalDataF, aes(x = UnemploymentSpell, colour = NameCluster)) +
    geom_density() +
    labs(title = 'Cluster Distribution',
         y     = 'Density',
         x     = 'Unemployment Spell') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05))) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm'),
          plot.margin = unit(c(0,0,0,0), 'lines')) +
    guides(colour = guide_legend('Cluster', override.aes = list(size = scale_factor * 1)))
  
  print(p)
  fileName = paste0('g3a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  insertImage(wb, 
              file = fileName, sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots,
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  
  # Looking at the nodes:
  p = finalDataF %>% mutate(node = factor(node)) %>% 
    ggplot(aes(x = UnemploymentSpell, colour = node)) +
    facet_wrap(NameCluster~.) +
    geom_density() +
    labs(title = 'Nodes Distribution',
         y     = 'Density',
         x     = 'Unemployment Spell') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05))) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(colour = guide_legend('Nodes', override.aes = list(size = scale_factor * 1)))
  
  print(p)
  fileName = paste0('g3b_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  insertImage(wb, 
              file = fileName, sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols,
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  
  rowLine = rowLine + jumpOfRows
  
  exportData = finalDataF %>% group_by(NameCluster) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = exportData %>% ggplot(aes(x = '', y = one, fill = NameCluster)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(title = 'Cluster Distribution', 
         y     = 'Percentage', 
         x     = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Cluster'))
  
  fileName = paste0('g4_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData, 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  rowLineSave = rowLine
  
  # Aggregates ------------------------------------------------------------
  finalDataFtoExport = finalDataF
  
  # * Experience ----------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Experience) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = exportData %>% ggplot(aes(x = Unity, y = one, fill = Experience)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black')+
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Experience'))
  
  print(p)
  fileName = paste0('g6a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Age -----------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Age) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Age)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black')+
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Age'))
  
  fileName = paste0('g7a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Governorate ---------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Governorate) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Governorate)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         x        = ' ', 
         y        = 'Percentage') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Governorate'))
  
  fileName = paste0('g8a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Disability ----------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Disability) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Disability)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Disability'))
  
  fileName = paste0('g9a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Gender --------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Gender) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Gender)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Gender'))
  
  fileName = paste0('g10a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Education Level -----------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, EducationLevel) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = EducationLevel)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Education level'))
  
  fileName = paste0('g11a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Unemployment Year ---------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, NameCluster) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = NameCluster)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Unemployment Year'))
  
  fileName = paste0('g12a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Same Job ------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, SameJob) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = SameJob)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Same Job'))
  
  fileName = paste0('g13a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Wage Group ----------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, WageGroup) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = WageGroup)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Wage Group'))
  
  fileName = paste0('g14a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% select(-c('Unity')), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Poverty Level -------------------------------------------------------
  p = ggplot(finalDataF, aes(x = Unity, y = Poverty)) +
    geom_boxplot() +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Poverty Level', 
         x        = ' ') +
    theme_classic() +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend(''))
  
  fileName = paste0('g15a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = as.data.frame(table(finalDataF$Unity, finalDataF$Poverty)), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Industry ------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Industry) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Industry)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Industry'))
  
  fileName = paste0('g16a_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData, 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # Individuals -----------------------------------------------------------
  rowLine = rowLineSave # Restores the initial position.
  
  # * Experience ----------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Experience) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Experience)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Experience'))
  
  fileName = paste0('g6_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Age -----------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Age) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Age)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Age'))
  
  fileName = paste0('g7_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Governorate ---------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Governorate) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Governorate)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Governorate'))
  
  fileName = paste0('g8_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Disability ----------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Disability) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Disability)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Disability'))
  
  fileName = paste0('g9_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Gender --------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Gender) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Gender)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Gender'))
  
  fileName = paste0('g10_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Education Level -----------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, EducationLevel) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = EducationLevel)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Education Level'))
  
  fileName = paste0('g11_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Unemployment Year ---------------------------------------------------
  # Same result as in aggregates...
  
  # p = ggplot(exportData, aes(x=UnemploymentYear, y=one, fill=NameCluster)) +
  #   geom_bar(stat='identity',position='fill')+
  #   labs(x=' ', y='Share', fill='Unemployment Year', subtitle = paste0(method,', ', gend, ', ', lastObs))
  # print(p)
  # fileName=paste0('g12_',gen,'_',method,'_',lastObs,'.png')
  # ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  # insertImage(wb, file = fileName, sheet = paste0(method,', ', gend, ', ', lastObs), startRow = rowLine, startCol = colPlots, width = 6 * widthExp, height = 4 * heightExp * widthExp)
  rowLine = rowLine + jumpOfRows
  
  # * Same Job ------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, SameJob) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = SameJob)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Same Job'))
  
  fileName = paste0('g13_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Wage Group ----------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, WageGroup) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = WageGroup)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Wage Group'))
  
  fileName = paste0('g14_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Poverty Level -------------------------------------------------------
  p = ggplot(finalDataF, aes(x = NameCluster, y = Poverty)) +
    geom_boxplot() +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Poverty Level', 
         x        = ' ') +
    theme_classic() +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend(''))
  
  fileName = paste0('g15_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = as.data.frame(table(finalDataF$NameCluster, finalDataF$Poverty)), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Industry ------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Industry) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Industry)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Industry'))
  
  fileName = paste0('g16_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # Next four plots:
  finalDataF$NameCluster                                = as.character(finalDataF$NameCluster)
  finalDataF$NameCluster[is.na(finalDataF$NameCluster)] = 'Missing'
  
  # * Average Unemployment Spell ------------------------------------------
  DataF                      = aggregate(finalDataF$UnemploymentSpell, 
                                         by  = list(finalDataF$NameCluster), 
                                         FUN = mean)
  colnames(DataF)            = c('Cluster', 'Unemployment Spell')
  DataF$`Unemployment Spell` = round(DataF$`Unemployment Spell`, 2)
  
  exportData = DataF %>% select(c('Cluster', 'Unemployment Spell'))
  p = ggplot(data = exportData, aes(x = Cluster, y = `Unemployment Spell`, fill = Cluster)) +
    geom_bar(stat = 'identity', position = position_dodge2(), colour = 'black', show.legend = F) +
    geom_text(aes(label = `Unemployment Spell`), vjust = 1.6, color = 'white', size = 3.5)+
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Average Unemployment Spell', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05))) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend(''))
  
  fileName = paste0('g17_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(Cluster, `Unemployment Spell`), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Standard Deviation Unemployment Spell -------------------------------
  DataF                      = aggregate(finalDataF$UnemploymentSpell, 
                                         by  = list(finalDataF$NameCluster), 
                                         FUN = sd)
  colnames(DataF)            = c('Cluster','Unemployment Spell')
  DataF$`Unemployment Spell` = round(DataF$`Unemployment Spell`, 2)
  
  exportData = DataF %>% select(c('Cluster', 'Unemployment Spell'))
  p = ggplot(data = exportData, aes(x = Cluster, y = `Unemployment Spell`, fill = Cluster)) +
    geom_bar(stat = 'identity', position = position_dodge2(), colour = 'black', show.legend = F) +
    geom_text(aes(label = `Unemployment Spell`), vjust = 1.6, color = 'white', size = 3.5)+
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Standard Deviation Unemployment Spell', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05))) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend(''))
  
  fileName = paste0('g18_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(Cluster, `Unemployment Spell`), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Variation Coefficient Unemployment Spell ----------------------------
  cv = function(x){
    return(sd(x)/mean(x))
  }
  
  DataF                      = aggregate(finalDataF$UnemploymentSpell, 
                                         by  = list(finalDataF$NameCluster), 
                                         FUN = cv)
  colnames(DataF)            = c('Cluster','Unemployment Spell')
  DataF$`Unemployment Spell` = round(DataF$`Unemployment Spell`, 2)
  
  exportData = DataF %>% select(c('Cluster', 'Unemployment Spell'))
  p = ggplot(data = exportData, aes(x = Cluster, y = `Unemployment Spell`, fill = Cluster)) +
    geom_bar(stat = 'identity', position = position_dodge2(), colour = 'black', show.legend = F) +
    geom_text(aes(label = `Unemployment Spell`), vjust = 1.6, color = 'white', size = 3.5)+
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Variation Coefficient Unemployment Spell', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05))) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend(''))
  
  fileName = paste0('g19_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(Cluster, `Unemployment Spell`), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Individuals Unemployment Spell --------------------------------------
  DataF                      = aggregate(finalDataF$UnemploymentSpell, 
                                         by  = list(finalDataF$NameCluster), 
                                         FUN = length)
  colnames(DataF)            = c('Cluster','Unemployment Spell')
  DataF$`Unemployment Spell` = round(DataF$`Unemployment Spell`, 2)
  
  exportData = DataF %>% select(c('Cluster', 'Unemployment Spell'))
  p = ggplot(data = exportData, aes(x = Cluster, y = `Unemployment Spell`, fill = Cluster)) +
    geom_bar(stat = 'identity', position = position_dodge2(), colour = 'black', show.legend = F) +
    geom_text(aes(label = `Unemployment Spell`), vjust = 1.6, color = 'white', size = 3.5)+
    labs(subtitle = paste0(method,', ', gend, ', ', lastObs),
         y        = 'Individuals Unemployment Spell', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::number_format(accuracy = 1, big.mark = ',')) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend(''))
  
  fileName = paste0('g20_',gen,'_',method,'_',lastObs,'.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs), 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData %>% spread(Cluster, `Unemployment Spell`), 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Distribution nodes --------------------------------------------------
  finalDataF$node = factor(finalDataF$node)
  exportData = finalDataF %>% group_by(NameCluster, node) %>% summarise(one = sum(one)) 
  p = exportData %>% ggplot(aes(x = NameCluster, y = node, fill = one)) +
    geom_tile(colour = 'black') +
    labs(title = 'Node Distribution', 
         y     = 'Nodes', 
         x     = 'Cluster',
         fill  = 'Cases') +
    theme_classic() +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Cluster'))
  
  fileName = paste0('g3_',gend,'_',method,'_',lastObs,'_','Nodes','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb, 
              file = fileName, 
              sheet = paste0(method,', ', gend, ', ', lastObs),
              startRow = rowLine, startCol = 1, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb, 
                 sheet    = paste0(method,', ', gend, ', ', lastObs), 
                 x        = exportData, 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows  
  
  
  # Final step ------------------------------------------------------------
  return(list(wb                   = wb,
              tr_FirstJob          = tr_FirstJob, 
              tr_Experience        = tr_Experience, 
              tr_Age               = tr_Age,  
              tr_Governorate       = tr_Governorate, 
              tr_Disability        = tr_Disability, 
              # tr_Gender            = tr_Gender, 
              tr_EducationLevel    = tr_EducationLevel, 
              tr_Industry          = tr_Industry, 
              tr_UnemploymentSpell = tr_UnemploymentSpell, 
              sCat                 = sCat,
              finalDataF           = finalDataFtoExport,
              distributionNodesPlot = p))
}

forecastingFunction <- function(gend, method, lastObs, clusterNum, wb2, tb) {
  # Preparation -------------------------------------------------------------
  dataN = read_dta('Input/NAF Clean Set.dta')
  colnames(dataN)
  # Creation of variables ---------------------------------------------------
  # - Dummy for first job:
  dataN$FirstJob = dataN$FirstJob_NAF
  dataN$FirstJob = factor(dataN$FirstJob,
                          levels = attr(tb$tr_Experience, 'xlevels')$FirstJob)
  
  # - Experience:
  dataN$Experience = round(dataN$Experience_NAF/365, 0)
  # dataN = dataN %>% mutate(Experience = case_when(Experience == 0 & (employment_status %in% c("Unemployed")) ~  0,
  #                                                 Experience == 0 & (!employment_status %in% c("Unemployed")) ~  1,
  #                                                 Experience == 1                    ~  1,
  #                                                 Experience >  1 & Experience <=  5 ~  5,
  #                                                 Experience >  5 & Experience <= 10 ~ 10,
  #                                                 Experience > 10 & Experience <= 15 ~ 15,
  #                                                 Experience > 15                    ~ 20))
  dataN = dataN %>% mutate(Experience = case_when(Experience == 0                    ~  0,
                                                  Experience == 1                    ~  1,
                                                  Experience >  1 & Experience <=  5 ~  5,
                                                  Experience >  5 & Experience <= 10 ~ 10,
                                                  Experience > 10 & Experience <= 15 ~ 15,
                                                  Experience > 15                    ~ 20))
  
  dataN$Experience = factor(dataN$Experience, 
                            levels = attr(tb$tr_FirstJob, 'xlevels')$Experience)
  
  # - Age:
  dataN$Age = factor(10 * round(dataN$Age_NAF/10, 0))
  dataN$Age = factor(dataN$Age, 
                     levels = attr(tb$tr_Experience, 'xlevels')$Age) # Measured in decades.
  
  # - Governorate:
  dataN$Governorate = factor(dataN$Gov_NAF, 
                             levels = attr(tb$tr_Experience, 'xlevels')$Governorate)
  
  # - Disability:
  dataN$Disability = factor(dataN$Disability_NAF, 
                            levels = attr(tb$tr_Experience, 'xlevels')$Disability)
  
  # - Gender:
  dataN$Gender = factor(dataN$Gender, 
                        levels = attr(tb$tr_Experience, 'xlevels')$Gender)
  
  # - Education Level:
  
  dataN = dataN %>% mutate(Educ_NAF = case_when(Educ_NAF %in% c("High Diploma", "Bachelor", "Masters", "PhD") ~  "Bachelor or above",
                                                Educ_NAF %in% c("Vocational Training") ~  "Vocational Training",
                                                Educ_NAF %in% c("Middle Diploma") ~  "Middle Diploma",
                                                Educ_NAF %in% c("Secondary or Below")~ "Secondary or Below" ))
  
  dataN$EducationLevel = factor(dataN$Educ_NAF, 
                                levels = attr(tb$tr_Experience, 'xlevels')$EducationLevel)
  
  # - Industry:
  dataN$Industry = factor(dataN$Industry_occupation_NAF, 
                          levels = attr(tb$tr_Experience, 'xlevels')$Industry)
  
  # Selecting the columns used...
  dataN=dataN%>%filter(Gender == gend) 
  #dataN = dataN %>% dplyr::select(FirstJob, Experience, Age, Governorate, Disability, Gender, EducationLevel, Industry) %>%
  #  filter(Gender == gend) 
  
  # Filling missing values --------------------------------------------------
  # temp                                                   = predict(tb$tr_FirstJob, newdata = dataN, type = 'class')
  # dataN[is.na(dataN$FirstJob),c('FirstJob')]             = temp[is.na(dataN$FirstJob)]
  temp                                                   = predict(tb$tr_Experience, dataN, type='class')
  dataN[is.na(dataN$Experience),c('Experience')]         = temp[is.na(dataN$Experience)]
  temp                                                   = predict(tb$tr_Age, dataN, type='class')
  dataN[is.na(dataN$Age),c('Age')]                       = temp[is.na(dataN$Age)]
  temp                                                   = predict(tb$tr_Governorate, dataN, type='class')
  dataN[is.na(dataN$Governorate),c('Governorate')]       = temp[is.na(dataN$Governorate)]
  temp                                                   = predict(tb$tr_Disability, dataN, type='class')
  dataN[is.na(dataN$Disability),c('Disability')]         = temp[is.na(dataN$Disability)]
  # temp                                                   = predict(tb$tr_Gender, dataN, type='class')
  # dataN[is.na(dataN$Gender),c('Gender')]                 = temp[is.na(dataN$Gender)] # Then this one shouldn't be necessary...
  temp                                                   = predict(tb$tr_EducationLevel, dataN, type='class')
  dataN[is.na(dataN$EducationLevel),c('EducationLevel')] = temp[is.na(dataN$EducationLevel)]
  temp                                                   = predict(tb$tr_Industry, dataN, type='class')
  dataN[is.na(dataN$Industry),c('Industry')]             = temp[is.na(dataN$Industry)]
  
  # Filtering ---------------------------------------------------------------
  # Filter according to the model (in this case remove Male)
  # Then check the prediction.
  rp = tb$tr_UnemploymentSpell
  newdata = dataN
  type = 'where'
  rp$frame$yval <- 1:nrow(rp$frame)
  should.be.leaves <- which(rp$frame[, 1] == "<leaf>")
  
  leaves <- predict(rp, newdata = newdata, type = "vector")
  should.be.leaves <- which (rp$frame[,1] == "<leaf>")
  bad.leaves <- leaves[!is.element (leaves, should.be.leaves)]
  if (length (bad.leaves) > 0){
    u.bad.leaves <- unique (bad.leaves)
    u.bad.nodes <- row.names (rp$frame)[u.bad.leaves]
    all.nodes <- row.names (rp$frame)[rp$frame[,1] == "<leaf>"]
    is.descendant <- function (all.leaves, node) {
      if (length (all.leaves) == 0) return (logical (0))
      all.leaves <- as.numeric (all.leaves); node <- as.numeric (node)
      if (missing (node)) return (NA)
      result <- rep (FALSE, length (all.leaves))
      for (i in 1:length (all.leaves)) {
        result=node<all.leaves
      }
      return (result)
    }
    where.tbl <- table (rp$where)
    names (where.tbl) <- row.names (rp$frame)[as.numeric (names (where.tbl))]
    for (u in 1:length (u.bad.nodes)) {
      desc.vec <- is.descendant (all.nodes, u.bad.nodes[u])
      me <- where.tbl[all.nodes][desc.vec]
      winner <- names (me)[me == max(me)][1]
      leaves[leaves == u.bad.leaves[u]] <- which (row.names (rp$frame) == winner)
    }
    leaves[leaves%in%u.bad.leaves] <- NA
  }
  
  s=leaves
  dataN$node=s
  dataN=merge(dataN, tb$sCat)
  
  # Be careful with the cluster number...
  dataN$NameCluster=factor(dataN$Cluster, levels = 1:clusterNum, labels=paste('Cluster', 1:clusterNum))
  dataN$Unity='Cluster'
  dataN$one=1
  
  # Exporting ---------------------------------------------------------------
  sheetName=paste0(method, ', ', gend,', ',lastObs)
  newSheet      = addWorksheet(wb2, sheetName = sheetName, gridLines = F)
  
  # - Excel:
  rowLine    = 1
  jumpOfRows = 30
  jumpOfCols = 20
  colPlots   = 1
  colTables  = 10
  # - By plot:
  base_exp     = 1
  heightExp    = 1
  widthExp     = 1
  scale_factor = base_exp/widthExp
  
  # 1| First column -------------------------------------------------------
  # Individuals -----------------------------------------------------------
  # finalDataF = tb$finalDataF
  finalDataF = dataN
  
  # Distribution clusters -------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster) %>% summarise(one = sum(one)) 
  p = exportData %>% ggplot(aes(x = '', y = one, fill = NameCluster)) +
    geom_bar(stat = 'identity', position = position_dodge2(), colour = 'black') +
    labs(title = 'Cluster Distribution', 
         y     = 'Number of individuals', 
         x     = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 1, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Cluster'))
  
  fileName = paste0('g4_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData, 
                 startRow = rowLine, startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Experience ----------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Experience) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Experience)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Experience'))
  
  fileName = paste0('g6_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, Experience) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Age -----------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Age) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Age)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Age'))
  
  fileName = paste0('g7_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, Age) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Governorate ---------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Governorate) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Governorate)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Governorate'))
  
  fileName = paste0('g8_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, Governorate) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Disability ----------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Disability) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Disability)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Disability'))
  
  fileName = paste0('g9_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, Disability) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Gender --------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Gender) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Gender)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Gender'))
  
  fileName = paste0('g10_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, Gender) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Education Level -----------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, EducationLevel) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = EducationLevel)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Education Level'))
  
  fileName = paste0('g11_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, EducationLevel) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # * Industry ------------------------------------------------------------
  exportData = finalDataF %>% group_by(NameCluster, Industry) %>% summarise(one = sum(one)) %>% 
    mutate(one = one/sum(one)*100)
  p = ggplot(exportData, aes(x = NameCluster, y = one, fill = Industry)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Industry'))
  
  fileName = paste0('g16_',gend,'_',method,'_',lastObs,'_','comparison (Individual)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% spread(NameCluster, one), 
                 startRow = rowLine, startCol = colTables)
  writeDataTable(wb2,
                 sheet    = sheetName,
                 x        = finalDataF %>% group_by(NameCluster, Industry) %>% summarise(one = sum(one)) %>% spread(NameCluster, one),
                 startRow = rowLine + nrow(exportData %>% spread(NameCluster, one)) + 2,
                 startCol = colTables)
  rowLine = rowLine + jumpOfRows
  
  # 2| Second column ------------------------------------------------------
  # Aggregates ------------------------------------------------------------
  rowLine = 1 + jumpOfRows
  finalDataFtoExport = tb$finalDataF
  
  # * Experience ----------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Experience) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, Experience) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = exportData %>% ggplot(aes(x = Unity, y = one, fill = Experience)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Experience'))
  
  print(p)
  fileName = paste0('g6a_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Age -----------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Age) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, Age) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Age)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Age'))
  
  fileName = paste0('g7a_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Governorate ---------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Governorate) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, Governorate) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Governorate)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         x        = ' ', 
         y        = 'Percentage') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Governorate'))
  
  fileName = paste0('g8a_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Disability ----------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Disability) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, Disability) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Disability)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Disability'))
  
  fileName = paste0('g9a_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Gender --------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Gender) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, Gender) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Gender)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic() +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Gender'))
  
  fileName = paste0('g10a_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Education Level -----------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, EducationLevel) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, EducationLevel) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = ggplot(exportData, aes(x = Unity, y = one, fill = EducationLevel)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic()   +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Education level'))
  
  fileName = paste0('g11ah_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Industry ------------------------------------------------------------
  exportData = finalDataF %>% group_by(Unity, Industry) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Evaluation')
  dataToAdd  = finalDataFtoExport %>% group_by(Unity, Industry) %>% summarise(one = sum(one)) %>% 
    ungroup() %>% mutate(one = one/sum(one)*100) %>% add_column(case = 'Training')
  exportData = bind_rows(dataToAdd, exportData) %>% mutate(case = factor(case, levels = c('Training', 'Evaluation')))
  p = ggplot(exportData, aes(x = Unity, y = one, fill = Industry)) +
    geom_bar(stat = 'identity', position = position_fill(), colour = 'black') +
    facet_wrap(case~.) +
    labs(subtitle = sheetName,
         y        = 'Percentage', 
         x        = ' ') +
    theme_classic()   +
    scale_y_continuous(expand = expansion(mult = c(0, .05)),
                       labels = scales::percent_format(scale = 100, suffix = '', accuracy = 1)) +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Industry'))
  
  fileName = paste0('g11a_',gend,'_',method,'_',lastObs,'_','comparison (Agreggate)','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData %>% select(-c('Unity')) %>% spread(case, one), 
                 startRow = rowLine, startCol = colTables + jumpOfCols)
  rowLine = rowLine + jumpOfRows
  
  # * Distribution nodes --------------------------------------------------
  finalDataF$node = factor(finalDataF$node)
  exportData = finalDataF %>% group_by(NameCluster, node) %>% summarise(one = sum(one)) 
  p = exportData %>% ggplot(aes(x = NameCluster, y = node, fill = one)) +
    geom_tile(colour = 'black') +
    labs(title = 'Node Distribution', 
         y     = 'Nodes', 
         x     = 'Cluster',
         fill  = 'Cases') +
    theme_classic() +
    theme(legend.position = 'right',
          text            = element_text(family = 'Georgia'),
          axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
          axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
          axis.ticks.x    = element_blank(),
          legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
          strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
          plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
          plot.subtitle   = element_text(size = 10, family = 'Georgia'),
          legend.key.size = unit(0.5 * scale_factor, 'cm')) +
    guides(fill   = guide_legend('Cluster'))
  
  fileName = paste0('g4_',gend,'_',method,'_',lastObs,'_','Nodes','.png')
  ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
  print(p)
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = 1, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  writeDataTable(wb2, 
                 sheet    = sheetName, 
                 x        = exportData, 
                 startRow = rowLine, startCol = colTables)
  fileName = paste0('g3_',gend,'_',method,'_',lastObs,'_','Nodes','.png')
  insertImage(wb2, 
              file = fileName, 
              sheet = sheetName, 
              startRow = rowLine, startCol = colPlots + jumpOfCols, 
              width = 6 * widthExp, height = 4 * heightExp * widthExp)
  rowLine = rowLine + jumpOfRows  
  
  ###Rounds
  Exp=attr(tb$tr_FirstJob, 'xlevels')$Experience
  Ag= attr(tb$tr_Experience, 'xlevels')$Age
  Gov=attr(tb$tr_Experience, 'xlevels')$Governorate
  Disa=attr(tb$tr_Experience, 'xlevels')$Disability
  Edu= c("Secondary or Below", "Vocational Training", "Middle Diploma", "Bachelor or above")
  Indu=attr(tb$tr_Experience, 'xlevels')$Industry
  
  
  
  
  Experience=matrix(0,0,0)
  Age=matrix(0,0,0)
  Governorate=matrix(0,0,0)
  Disability=matrix(0,0,0)
  Gender=matrix(0,0,0)
  EducationLevel=matrix(0,0,0)
  
  
  
  
  for(b in Ag){
    for(c in Gov){
      for(d in Disa){
        for(f in Edu){
          for(a in Exp){
            Experience=c(Experience,a)
            Age=c(Age,b)
            Governorate=c(Governorate,c)
            Disability=c(Disability,d)
            EducationLevel=c(EducationLevel,f)
          } 
        } 
      } 
    } 
  }
  transitionMat=data.frame(Experience=factor(Experience, levels =Exp ),
                           Age=factor(Age, levels =Ag ),
                           Governorate=factor(Governorate, levels =Gov ),
                           Disability=factor(Disability, levels =Disa ),
                           EducationLevel=factor(EducationLevel, levels =Edu),
                           Gender=factor(gend, levels =c("Male", "Female"))
  )
  
  rp = tb$tr_UnemploymentSpell
  newdata = transitionMat
  type = 'where'
  rp$frame$yval <- 1:nrow(rp$frame)
  should.be.leaves <- which(rp$frame[, 1] == "<leaf>")
  
  leaves <- predict(rp, newdata = newdata, type = "vector")
  should.be.leaves <- which (rp$frame[,1] == "<leaf>")
  bad.leaves <- leaves[!is.element (leaves, should.be.leaves)]
  if (length (bad.leaves) > 0){
    u.bad.leaves <- unique (bad.leaves)
    u.bad.nodes <- row.names (rp$frame)[u.bad.leaves]
    all.nodes <- row.names (rp$frame)[rp$frame[,1] == "<leaf>"]
    is.descendant <- function (all.leaves, node) {
      if (length (all.leaves) == 0) return (logical (0))
      all.leaves <- as.numeric (all.leaves); node <- as.numeric (node)
      if (missing (node)) return (NA)
      result <- rep (FALSE, length (all.leaves))
      for (i in 1:length (all.leaves)) {
        result=node<all.leaves
      }
      return (result)
    }
    where.tbl <- table (rp$where)
    names (where.tbl) <- row.names (rp$frame)[as.numeric (names (where.tbl))]
    for (u in 1:length (u.bad.nodes)) {
      desc.vec <- is.descendant (all.nodes, u.bad.nodes[u])
      me <- where.tbl[all.nodes][desc.vec]
      winner <- names (me)[me == max(me)][1]
      leaves[leaves == u.bad.leaves[u]] <- which (row.names (rp$frame) == winner)
    }
    leaves[leaves%in%u.bad.leaves] <- NA
  }
  
  s=leaves
  transitionMat$node=s
  transitionMat=merge(transitionMat, tb$sCat, all.x=TRUE)
  transitionMat$Cluster =factor(transitionMat$Cluster)
  transitionMat$Cases=1
  
  
  for(b in Ag[c(3:6)]){
    for(c in Gov[c(3,5,7,13)]){
      for(d in Disa[1]){
        g=transitionMat%>%filter(Governorate==c, Disability==d, Age==b)%>%
          ggplot(aes(fill = Cluster, x =  Experience, y = EducationLevel)) +
          geom_tile()+
          labs(title = paste("Age:",b,"| Governorate:",c, "| Disability:",d),
               x     = 'Experience', 
               y     = 'Education Level',
               fill  = 'Cluster') +
          theme_classic() +
          theme(legend.position = 'right',
                text            = element_text(family = 'Georgia'),
                axis.text.x     = element_text(size = scale_factor * 10, family = 'Georgia'),
                axis.text.y     = element_text(size = scale_factor * 10, family = 'Georgia'),
                axis.title.x    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
                axis.title.y    = element_text(hjust = 0.5, size = scale_factor * 10, family = 'Georgia'),
                axis.ticks.x    = element_blank(),
                legend.text     = element_text(size = scale_factor * 10, family = 'Georgia'),
                strip.text      = element_text(size = scale_factor * 10, family = 'Georgia'),
                plot.title      = element_text(face = 'bold', size = 10, family = 'Georgia'),
                plot.subtitle   = element_text(size = 10, family = 'Georgia'),
                legend.key.size = unit(0.5 * scale_factor, 'cm')) +
          guides(fill   = guide_legend('Cluster'))
        
        fileName = paste0(gend,method,lastObs,b,c,d,'.png')
        ggsave(fileName, width = 6 * widthExp, height = 4 * heightExp * widthExp, scale = scale_factor)
        print(g)
        insertImage(wb2, 
                    file = fileName, 
                    sheet = sheetName, 
                    startRow = rowLine, startCol = 1, 
                    width = 6 * widthExp, height = 4 * heightExp * widthExp)
        rowLine = rowLine + jumpOfRows  
      } 
    } 
  } 
  
  sheetName2=paste0("List", method, ', ', gend,', ',lastObs)
  newSheet2      = addWorksheet(wb2, sheetName = sheetName2, gridLines = F)
  dataT=dataN%>%select(c("Request_ID", "NationalID_Number", "Head_Of_Family_National_Number","Family_Member_Name","Head_Of_Family_Name","Family_Member_Birth_Date","Gender","Age_NAF","Educ_NAF","Gov_NAF","Profession_Type_Name", "Type_of_Technical_Training_Desc", "Telephonenumber","employment_status","poverty_score","no_member_HH","Cluster"))
  writeDataTable(wb2, 
                 sheet    = sheetName2, 
                 x        = dataT, 
                 startRow = 1, startCol = 1)
  
  # Final step ------------------------------------------------------------
  return(wb2)
}

# 2| Unemployment profiling -----------------------------------------------
# 2.1| Three clusters -----------------------------------------------------
startTime  = Sys.time()
gend       = 'Male'
method     = 'Method 3'
lastObs    = 'One'
clusterNum = 3

wb  = openxlsx::createWorkbook()
wb2 = openxlsx::createWorkbook()
# newTemp = grapherFunction('Female', 'Method 2', 'One', 3, wb)


for(i in c('Male', "Female")){ # c('Male', 'Female')
  for(j in c( 'Method 3')){ # 'Method 2',
    for(k in c('One')){ # 'All', 
      gend       = i
      method     = j
      lastObs    = k
      newTemp = grapherFunction(gend, method, lastObs, clusterNum, wb)
      wb      = newTemp$wb
      tb  = list(wb                   = wb,
                 tr_FirstJob          = newTemp$tr_FirstJob, 
                 tr_Experience        = newTemp$tr_Experience, 
                 tr_Age               = newTemp$tr_Age,  
                 tr_Governorate       = newTemp$tr_Governorate, 
                 tr_Disability        = newTemp$tr_Disability, 
                 tr_EducationLevel    = newTemp$tr_EducationLevel, 
                 tr_Industry          = newTemp$tr_Industry, 
                 tr_UnemploymentSpell = newTemp$tr_UnemploymentSpell, 
                 sCat                 = newTemp$sCat,
                 finalDataF           = newTemp$finalDataF,
                 distributionNodesPlot=newTemp$p)
      
      
      wb2 = forecastingFunction(gend, method, lastObs, clusterNum, wb2, tb)
    }
  }
}

saveWorkbook(wb, 
             file = paste0('Unemployment Profiling (', clusterNum, ')', '.xlsx'), 
             overwrite = TRUE)
saveWorkbook(wb2, 
             file = paste0('NAF Analysis Daily  (', clusterNum, ')', '.xlsx'), 
             overwrite = TRUE)

endTime = Sys.time()
endTime - startTime
