# Converts the digitized data for Rosario de Santa Fe 1886-1900 into the
# Station Exchange Format.
#
# Requires file write_sef.R and libraries XLConnect, plyr
#
# Created by Yuri Brugnara, University of Bern - 20 Dec 2018

###############################################################################


require(XLConnect)
require(plyr)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


infile <- "../data/raw/RosarioDeSantaFe/Rosario_De_Santa_Fe_Argentina_May1886-Dec1900.xls"
outpath <- "../data/formatted/"

lon <- -60.333
lat <- -32.945
alt <- 36

variables <- c("ta", "p", "Tx", "Tn", "tb", "dd", 
               "n", "rr", "wind_force", "ss", "Ts")
units <- c("C", "Pa", "C", "C", "C", "degree", "%", "mm", "", "hours", "C")
time_flags <- c(0, 0, 13, 13, 0, 0, 0, 12, 0, 12, 0)
conversions <- list(ta = function(x) round(x, 1),
                    p = function(x, y) 
                      round(100 * convert_pressure(x, f = 1, lat = lat, 
                                                   alt = alt, atb = y), 0),
                    Tx = function(x) round(x, 1),
                    Tn = function(x) round(x, 1),
                    tb = function(x) round(x, 1),
                    dd = function(x) round(x, 0),
                    n = function(x) round(10 * x, 0),
                    rr = function(x) round(x, 1),
                    wind_force = function(x) round(x, 0),
                    ss = function(x) round(x, 2),
                    Ts = function(x) round(x, 1))


read_template <- function(startRow, endRow,
                          times, # character vector of times
                          endTimes, # endCol of each time (vector)
                          vars, # list of vectors with variable codes (one for each time)
                          file = infile) {
  
  all_vars <- unlist(vars)
  data_types <- rep("numeric", length(all_vars) + 3)
  data_types[which(all_vars %in% c("dd", NA)) + 3] <- "character"
  template <- readWorksheetFromFile(file, sheet = 2,
                                    startRow = startRow,
                                    endRow = endRow,
                                    header = FALSE,
                                    colTypes = data_types,
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
    ## Add one day to the evening observations
    if (i == 3) {
      dates <- paste(tmp[[i]]$y, tmp[[i]]$m, tmp[[i]]$d, sep = "-")
      times <- strptime(dates, format = "%Y-%m-%d") + 3600 * 24 # 
      tmp[[i]]$y <- as.integer(format(times, "%Y"))
      tmp[[i]]$m <- as.integer(format(times, "%m"))
      tmp[[i]]$d <- as.integer(format(times, "%d"))
    }
  }
  template <- rbind.fill(tmp)
  
  return(template)
  
}


## Initialize data frames
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
template1 <- read_template(6, 233, c("1217", "2217"), c(17, 31),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA)))
template1$rr_orig <- paste0("Orig=", round(template1$rr, 2), "in")
template1$rr <- 25.4 * template1$rr

## Read data (2nd part - 3 observations per day - precipitation in inches)
template2 <- read_template(242, 667, c("1117", "1817", "0117"), c(19, 35, 51),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA, NA, NA)))
template2$rr_orig <- paste0("Orig=", round(template2$rr, 2), "in")
template2$rr <- 25.4 * template2$rr

## Read data (3rd part - 3 observations per day - precipitation in mm)
template3 <- read_template(1038, 1583, c("1117", "1817", "0117"), c(17, 29, 42),
                           list(c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", NA, NA),
                                c("atb", "p", "ta", "tb", "n", NA, NA, "dd", 
                                  "wind_force", "rr", NA, NA, "ss")))
template3$rr_orig <- ""

## Read data (4th part - 2 observations per day - precipitation in mm)
template4 <- read_template(1591, 3985, c("1217", "2217"), c(24, 45),
                           list(c("atb", "p", "ta", "tb", NA, "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", "Ts", NA, NA,
                                  NA, NA, NA, NA, NA),
                                c("atb", "p", "ta", "tb", NA, "n", NA, NA, "dd", 
                                  "wind_force", "rr", "Tx", "Tn", NA, NA, NA, NA,
                                  "ss", NA, NA, NA)))
template4$rr_orig <- ""


## Merge templates
template_all <- rbind.fill(template1, template2, template3, template4)
template_all <- template_all[which(!is.na(template_all$y)), ]
template_all$p_orig <- paste0("Orig=", round(template_all$p, 1), "mm,atb=", 
                              round(template_all$atb, 1), "C")
template_all$dd_orig <- paste0("Orig=", template_all$dd)
template_all$n_orig <- paste0("Orig=", round(template_all$n, 0))
template_all[, paste(c("ta", "tb", "Tx", "Tn", "wind_force", "ss", "Ts"), 
                     "orig", sep = "_")] <- ""


## Transform wind direction to degrees
directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
template_all$dd <- 22.5 * (match(toupper(template_all$dd), directions) - 1)


## Set missing time to those days marked by asterisks
template_all$h[which(is.na(template_all$d))] <- ""
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


## Write output
for (i in 1:length(variables)) {
  ## First remove missing values, order by time and add column with variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  Data[[variables[i]]] <- Data[[variables[i]]][order(Data[[variables[i]]]$y,
                                                     Data[[variables[i]]]$m,
                                                     Data[[variables[i]]]$d,
                                                     Data[[variables[i]]]$h), ]
  Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]])
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            cod = "Rosario_Santa_Fe",
            nam = "Rosario de Santa Fe",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAmerica",
            repo = "",
            units = units[i],
            metaHead = ifelse(variables[i] == "p", "PTC=T,PGC=T", ""),
            meta = Data[[variables[i]]][, 7],
            timef = time_flags[i])
}