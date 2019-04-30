# Converts the Daily Weather Reports digitized data for 1902 into the
# Station Exchange Format.
#
# Requires libraries XLConnect and dataresqc
# (https://github.com/c3s-data-rescue-service/dataresqc)
#
# Created by Yuri Brugnara, University of Bern - 30 Apr 2019

###############################################################################


require(XLConnect)
require(dataresqc)
options(scipen = 999) # avoid exponential notation


inpath <- "../data/raw/DWR/"
listfile <- paste0(inpath, "List_of_stations.xlsx")
outpath <- "../data/formatted/"

# Define variable codes and units
variables <- c("ta", "mslp", "Tx", "Tn", "rh", "dd", "n", "wind_force")
units <- c("C", "Pa", "C", "C", "%", "degree", "%", "")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round(x, 1),
                    mslp = function(x) round(100 * convert_pressure(x, f = 1), 0),
                    Tx = function(x) round(x, 1),
                    Tn = function(x) round(x, 1),
                    rh = function(x) round(x, 0),
                    dd = function(x) round(x, 0),
                    n = function(x) round(33.3 * x, 0),
                    wind_force = function(x) round(x, 0))

# Define function to get rid of latin characters
convert_to_ascii <- function(x) {
  x <- gsub("ñ", "n", x)
  x <- gsub("á", "a", x)
  x <- gsub("é", "e", x)
  x <- gsub("í", "i", x)
  x <- gsub("ó", "o", x)
  x <- gsub("ú", "u", x)
}


## Read station names
stations <- readWorksheetFromFile(listfile, sheet = 1, startRow = 2, header = FALSE)[, 1]


## Loop on stations
for (ist in 1:length(stations)) {
  print(stations[ist])
  
  
  ## Initialize data frames
  Data <- list()
  for (v in variables) {
    Data[[v]] <- data.frame(year = integer(),
                            month = integer(),
                            day = integer(),
                            hour = integer(),
                            minute = integer(),
                            value = numeric(),
                            orig = character())
  }
  
  
  ## Read data files
  infile <- convert_to_ascii(paste0(inpath, 
                                    gsub(" ", "_", stations[ist]), ".xlsx"))
  template <- readWorksheetFromFile(infile, sheet = 1, startRow = 5,
                                    header = FALSE, endCol = 16,
                                    colTypes = c(rep("numeric", 3),
                                                 rep("character", 2),
                                                 rep("numeric", 7),
                                                 "character",
                                                 "numeric",
                                                 "character",
                                                 "numeric"),
                                    drop = c(4, 5, 7, 9, 15),
                                    forceConversion = TRUE,
                                    autofitCol = FALSE,
                                    readStrategy = "fast")
  names(template) <- c("y", "m", "d", "mslp", "ta", "Tx", "Tn", 
                       "rh", "dd", "wind_force", "n")
  template <- template[which(!is.na(template$y)), ]
  template$dd_orig <- paste0("Orig=", template$dd)
  template$n_orig <- paste0("Orig=", round(template$n, 0))
  template$mslp_orig <- paste0("Orig=", round(template$mslp, 1), "mm")
  template[, paste(c("ta", "Tx", "Tn", "rh", "wind_force"), "orig", sep = "_")] <- ""
  
  
  ## Time is 2 PM LST until August and 7 AM LST from September
  template$HH <- NA
  template$HH[which(template$m <= 8)] <- 14
  template$HH[which(template$m > 8)] <- 7
  template$MM <- 0
  
  
  ## Transform wind direction to degrees
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(template$dd), directions) - 1)
  
  
  ## Write to data frames
  for (i in 1:length(variables)) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c(names(template)[1:3], "HH", "MM", 
                                               variables[i], 
                                               paste0(variables[i], "_orig"))])
  }
  
  
  ## Output
  for (i in 1:length(variables)) {
    
    ## Remove missing values
    Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
    
    ## Define time statistic
    if (variables[i] == "Tx") {
      tstat <- "maximum"
    } else if (variables[i] == "Tn") {
      tstat <- "minimum"
    } else tstat <- "point"
    
    ## Write file
    if (dim(Data[[variables[i]]])[1] > 0) {
      write_sef(Data = Data[[variables[i]]][, 1:6],
                outpath = outpath,
                variable = variables[i],
                cod = convert_to_ascii(paste0("DWR_", 
                                              substr(gsub(" ", "_", stations[ist]), 1, 12))),
                nam = stations[ist],
                lat = "",
                lon = "",
                alt = "",
                sou = "C3S_SouthAmerica",
                link = "",
                stat = tstat,
                units = units[i],
                metaHead = ifelse(variables[i] == "mslp", "PGC=F", ""),
                meta = Data[[variables[i]]][, 7],
                period = ifelse(variables[i] %in% c("Tx","Tn"), "day", 0),
                time_offset = -4.28)
    }
  }
  
}