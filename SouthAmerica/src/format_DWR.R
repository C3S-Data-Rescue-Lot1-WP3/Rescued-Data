# Converts the Daily Weather Reports digitized data for 1902 into the
# Station Exchange Format.
#
# Requires file write_sef.R and library XLConnect
#
# Created by Yuri Brugnara, University of Bern - 20 Dec 2018

###############################################################################


require(XLConnect)
source("write_sef.R")
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
                            time = character(),
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
  
  
  ## Time is 1817 UTC until August and 1117 UTC from September
  template$h <- NA
  template$h[which(template$m <= 8)] <- "1817"
  template$h[which(template$m > 8)] <- "1117"
  
  
  ## Transform wind direction to degrees
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(template$dd), directions) - 1)
  
  
  ## Write to data frames
  for (i in 1:length(variables)) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c(names(template)[1:3], "h", variables[i], 
                                               paste0(variables[i], "_orig"))],
                                  stringsAsFactors = FALSE)
  }
  
  
  ## Write output
  for (i in 1:length(variables)) {
    ## First remove missing values and add column with variable code (required by write_sef)
    Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
    if (dim(Data[[variables[i]]])[1] > 0) {
      Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]],
                                    stringsAsFactors = FALSE)
      write_sef(Data = Data[[variables[i]]][, 1:6],
                outpath = outpath,
                cod = convert_to_ascii(paste0("DWR_", 
                                              substr(gsub(" ", "_", stations[ist]), 1, 12))),
                nam = stations[ist],
                lat = "",
                lon = "",
                alt = "",
                sou = "C3S_SouthAmerica",
                repo = "",
                units = units[i],
                metaHead = "",
                meta = Data[[variables[i]]][, 7],
                timef = ifelse(variables[i] %in% c("Tx", "Tn"), 13, 0))
    }
  }
  
}