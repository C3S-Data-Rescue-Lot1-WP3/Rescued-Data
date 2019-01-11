# Converts the Uitenhage digitized data from the University of Witwatersrand 
# into the Station Exchange Format.
#
# Requires file write_sef.R and library XLConnect
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


lat <- -33.7687
lon <- 25.4141
alt <- 103

inpath <- "../data/raw/Uitenhage/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("ta", "p", "dd")
units <- c("C", "Pa", "degree")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(100 * convert_pressure(x, f = 25.4,
                                                   lat = lat, alt = alt), 0),
                    dd = function(x) round(x, 0))

# Define function to convert the month name into a number
get_month <- function(x) {
  x[grep("Jan", x, TRUE)] <- 1
  x[grep("Feb", x, TRUE)] <- 2
  x[grep("Mar", x, TRUE)] <- 3
  x[grep("Apr", x, TRUE)] <- 4
  x[grep("May", x, TRUE)] <- 5
  x[grep("Jun", x, TRUE)] <- 6
  x[grep("Jul", x, TRUE)] <- 7
  x[grep("Aug", x, TRUE)] <- 8
  x[grep("Sep", x, TRUE)] <- 9
  x[grep("Oct", x, TRUE)] <- 10
  x[grep("Nov", x, TRUE)] <- 11
  x[grep("Dec", x, TRUE)] <- 12
  return(x)
}


# Define function to fill in missing dates
fill <- function(x) {
  for (i in 2:length(x)) {
    if (is.na(x[i])) x[i] <- x[i - 1]
  }
  return(x)
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


## Read data files
infiles <- list.files(inpath)
for (infile in infiles) {
  year <- as.integer(substr(infile, 11, 14))
  
  ## Template 1 (Jan-Feb 1821)
  if (substr(infile, 17, 19) == "Jan") {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 10, endRow = 68,
                                      header = FALSE, sheet = 1, endCol = 5,
                                      colTypes = c("character",
                                                   rep("numeric", 3),
                                                   "character"),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "p", "ta", "dd")
    template$h <- ""
    template$p_orig <- paste0("Orig=", round(template$p, 2), "in")
    template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    
  ## Template 2 (Mar 1821 - 1824)
  } else {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12,
                                      header = FALSE, sheet = 1, endCol = 6,
                                      colTypes = c("character",
                                                   "numeric",
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character"),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "p", "ta", "dd")
    template$p_orig <- paste0("Orig=", round(template$p, 2), "in")
    template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    j <- which(!is.na(template$h))
    template[j, c("p_orig", "ta_orig", "dd_orig")] <-
      sapply(template[j, c("p_orig", "ta_orig", "dd_orig")],
             paste, paste0("t=", template$h[j]), sep = ",")
    template$h <- ""
  }
  
  template$y <- year
  template$m <- as.integer(fill(get_month(template$m)))
  template$d <- fill(template$d)
  template <- template[which(!is.na(template$m)), ]

    
  ## Transform wind direction to degrees
  ## Entries like 'NWbN' are converted as 'NW'
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(sapply(strsplit(template$dd, "b"), 
                                              function(x) x[1])), directions) - 1)

  ## Write to data frames
  for (i in 1:length(variables)) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c("y", "m", "d", "h", variables[i], 
                                               paste0(variables[i], "_orig"))])
  }
}


## Write output
for (i in 1:length(variables)) {
  ## First remove missing values and add variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]])
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            cod = "Uitenhage",
            nam = "Uitenhage",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            repo = "",
            units = units[i],
            metaHead = ifelse(i==2, "PTC=F,PGC=T", ""),
            meta = Data[[variables[i]]][, 7],
            timef = 0)
}