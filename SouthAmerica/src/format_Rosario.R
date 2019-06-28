# Converts the digitized data for Rosario de Santa Fe 1886-1900 into the
# Station Exchange Format v0.2.
#
# Requires libraries XLConnect, plyr, dataresqc
# (https://github.com/c3s-data-rescue-service/dataresqc)
#
# Created by Yuri Brugnara, University of Bern - 30 Apr 2019

###############################################################################


require(XLConnect)
require(plyr)
require(dataresqc)
options(scipen = 999) # avoid exponential notation


infile <- "../data/raw/RosarioDeSantaFe/Rosario_De_Santa_Fe_Argentina_May1886-Dec1900.xls"
outpath <- "../data/formatted/"

lon <- -60.333
lat <- -32.945
alt <- 35.7

variables <- c("ta", "p", "Tx", "Tn", "tb", "dd", 
               "n", "rr", "w", "ss", "Ts")
units <- c("C", "Pa", "C", "C", "C", "degree", "%", "mm", "", "hours", "C")
stat_flags <- c("point", "point", "maximum", "minimum", "point", "point", 
                "point", "sum", "point", "sum", "point")
varids <- c(6, 4, 0, 1, 7, 2, 3, 5, 8, 9, 10) # needed to build the link to the C3S registry

conversions <- list(ta = function(x) round(as.numeric(x), 1),
                    p = function(x, y) 
                      round(100 * convert_pressure(as.numeric(x), f = 1, lat = lat, 
                                                   alt = alt, atb = as.numeric(y)), 0),
                    Tx = function(x) round(as.numeric(x), 1),
                    Tn = function(x) round(as.numeric(x), 1),
                    tb = function(x) round(as.numeric(x), 1),
                    dd = function(x) round(as.numeric(x), 0),
                    n = function(x) round(10 * as.numeric(x), 0),
                    rr = function(x) round(as.numeric(x), 1),
                    w = function(x) round(as.numeric(x), 0),
                    ss = function(x) round(as.numeric(x), 2),
                    Ts = function(x) round(as.numeric(x), 1))


read_template <- function(startRow, endRow,
                          times, # character vector of observation times
                          endTimes, # endCol of each observation time (vector)
                          vars, # list of vectors with variable codes (one for each observation time)
                          file = infile) {
  
  all_vars <- unlist(vars)
  template <- readWorksheetFromFile(file, sheet = 2,
                                    startRow = startRow,
                                    endRow = endRow,
                                    endCol = rev(endTimes)[1],
                                    header = FALSE,
                                    colTypes = c(rep("numeric", 3), 
                                                 rep("character", length(all_vars))),
                                    forceConversion = TRUE,
                                    autofitCol = FALSE,
                                    readStrategy = "fast")
  
  tmp <- list()
  tmp[[1]] <- template[, 1:endTimes[1]]
  keep <- which(!is.na(vars[[1]]))
  tmp[[1]] <- tmp[[1]][, c(1:3, keep + 3)]
  names(tmp[[1]]) <- c("y", "m", "d", vars[[1]][keep])
  tmp[[1]]$h <- times[1]
  for (i in 2:length(times)) {
    tmp[[i]] <- template[, (endTimes[i-1]+1):endTimes[i]]
    keep <- which(!is.na(vars[[i]]))
    tmp[[i]] <- tmp[[i]][, keep]
    names(tmp[[i]]) <- vars[[i]][keep]
    tmp[[i]] <- cbind(tmp[[1]][, 1:3], tmp[[i]])
    tmp[[i]]$h <- times[i]
  }
  template <- rbind.fill(tmp)
  
  return(template)
  
}


## Initialize list of data frames (one per variable)
Data <- list()
for (v in variables) {
  Data[[v]] <- data.frame(year = integer(),
                          month = integer(),
                          day = integer(),
                          time = character(),
                          value = numeric(),
                          orig = character())
}


## Read data (1st part - 2 observations per day - precipitation in inches)
template1 <- read_template(6, 233, c("0800", "1800"), c(17, 31),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA)))
template1$rr_orig <- paste0("Orig=", template1$rr, "in")
template1$rr <- 25.4 * as.numeric(template1$rr)

## Read data (2nd part - 3 observations per day - precipitation in inches)
template2 <- read_template(242, 667, c("0700", "1400", "2100"), c(19, 35, 51),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA, NA, NA)))
template2$rr_orig <- paste0("Orig=", template2$rr, "in")
template2$rr <- 25.4 * as.numeric(template2$rr)

## Read data (3rd part - 3 observations per day - precipitation in mm)
template3 <- read_template(1038, 1402, c("0700", "1400", "2100"), c(17, 29, 42),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", NA, NA, "ss")))
template3$rr_orig <- paste0("Orig=", template3$rr, "mm")

## Read data (4th part - 3 observations per day - precipitation in mm, zeros omitted)
template4 <- read_template(1403, 1583, c("0700", "1400", "2100"), c(17, 29, 42),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "w", "rr", NA, NA, "ss")))
template4$rr[is.na(template4$rr)] <- 0
template4$rr_orig <- paste0("Orig=", template4$rr, "mm")

## Read data (5th part - 2 observations per day - precipitation in mm, zeros omitted)
template5 <- read_template(1591, 2921, c("0800", "1800"), c(24, 45),
                           list(c("atb", "p", "ta", "tb", NA, "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", "Ts", NA, NA,
                                  NA, NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", NA, "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA, NA, NA,
                                  "ss", NA, NA, NA)))
template5$rr[is.na(template5$rr)] <- 0
template5$rr_orig <- paste0("Orig=", template5$rr, "mm")

## Read data (6th part - 2 observations per day - precipitation in mm)
template6 <- read_template(2922, 3985, c("0800", "1800"), c(24, 45),
                           list(c("atb", "p", "ta", "tb", NA, "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", "Ts", NA, NA,
                                  NA, NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", NA, "n", NA, NA, "dd", 
                                  "w", "rr", "Tx", "Tn", NA, NA, NA, NA,
                                  "ss", NA, NA, NA)))
template6$rr_orig <- paste0("Orig=", template6$rr, "mm")


## Merge templates
template_all <- rbind.fill(template1, template2, template3, template4, template5, template6)
template_all <- template_all[which(!is.na(template_all$y)), ]
template_all$p_orig <- paste0("Orig=", template_all$p, "mm|atb=", 
                              template_all$atb, "C")
template_all$dd_orig <- paste0("Orig=", template_all$dd)
template_all$n_orig <- paste0("Orig=", template_all$n)
template_all[, paste(c("ta", "tb", "Tx", "Tn", "w", "ss", "Ts"), 
                     "orig", sep = "_")] <- ""


## Transform wind direction to degrees
directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
template_all$dd <- 22.5 * (match(toupper(template_all$dd), directions) - 1)


## Set missing time to those days marked by asterisks
template_all$h[which(is.na(template_all$d))] <- NA
template_all$d[which(is.na(template_all$d))] <- 
  template_all$d[which(is.na(template_all$d)) - 1] + 1
template_all$d[which(is.na(template_all$d))] <- 
  template_all$d[which(is.na(template_all$d)) - 1] + 1


## Write to data frames
for (i in 1:length(variables)) {
  if (variables[i] == "p") {
    template_all[, variables[i]] <- 
      conversions[[variables[i]]](template_all[, variables[i]], template_all$atb)
  } else {
    template_all[, variables[i]] <- 
      conversions[[variables[i]]](template_all[, variables[i]])
  }
  Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                template_all[, c(names(template_all)[1:3], "h", 
                                                 variables[i], 
                                                 paste0(variables[i], "_orig"))])
}


## Output
for (i in 1:length(variables)) {
  
  ## Order by time
  Data[[variables[i]]] <- Data[[variables[i]]][order(Data[[variables[i]]]$y,
                                                     Data[[variables[i]]]$m,
                                                     Data[[variables[i]]]$d,
                                                     Data[[variables[i]]]$h), ]
  
  ## Split time into hour and minute
  Data[[variables[i]]]$HH <- as.integer(substr(Data[[variables[i]]]$h, 1, 2))
  Data[[variables[i]]]$MM <- as.integer(substr(Data[[variables[i]]]$h, 3, 4))
  
  ## Add original time in meta column
  separator <- ifelse(Data[[variables[i]]][1,6] == "", "", "|")
  Data[[variables[i]]][, 6] <- 
    paste0(Data[[variables[i]]][, 6], separator, "orig.time=", Data[[variables[i]]][, 4])
  
  ## Write file
  write_sef(Data = Data[[variables[i]]][, c(1:3,7,8,5)],
            outpath = outpath,
            variable = variables[i],
            cod = "Rosario_Santa_Fe",
            nam = "Rosario de Santa Fe",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAmerica",
            link = paste0("https://data-rescue.copernicus-climate.eu/lso/", 1086326 + varids[i]),
            stat = stat_flags[i],
            units = units[i],
            metaHead = paste0("Data policy=GNU GPL v3.0", 
                             ifelse(variables[i] == "p", "|PTC=Y|PGC=Y", "")),
            meta = Data[[variables[i]]][, 6],
            period = ifelse(stat_flags[i]=="point", "0", "p1day"),
            time_offset = -4.29)
}