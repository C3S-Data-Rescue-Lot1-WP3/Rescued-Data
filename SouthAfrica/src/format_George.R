# Converts the George digitized data from the University of Witwatersrand 
# into the Station Exchange Format.
#
# Requires file write_sef.R and library XLConnect
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


lat <- -33.9881
lon <- 22.453
alt <- ""

inpath <- "../data/raw/George/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("ta", "p", "dd")
units <- c("C", "Pa", "degree")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(100 * convert_pressure(x, f = 25.4), 0),
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
for (ifile in 1:length(infiles)) {
  year <- as.integer(substr(strsplit(infiles[ifile], "_")[[1]][2], 1, 4))
  template <- readWorksheetFromFile(paste0(inpath, infiles[ifile]), 
                                    startRow = ifelse(year == 1821, 12, 11),
                                    header = FALSE, sheet = 1, endCol = 5,
                                    colTypes = c("character",
                                                 rep("numeric", 3),
                                                 "character"),
                                    forceConversion = TRUE,
                                    readStrategy = "fast")
  names(template) <- c("m", "d", "p", "ta", "dd")
  template$y <- year
  template$m <- as.integer(fill(get_month(template$m)))
  template$d <- fill(template$d)
  template$h <- ""
  template <- template[which(!is.na(template$m)), ]

  ## Organize meta column
  template$p_orig <- paste0("Orig=", round(template$p, 3), "in")
  template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")
  template$dd_orig <- paste0("Orig=", template$dd)
  
  ## Transform wind direction to degrees
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(template$dd), directions) - 1)

  ## Write to data frames
  for (i in 1:length(variables)) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c("y", "m", "d", "h", variables[i], 
                                               paste0(variables[i], "_orig"))],
                                  stringsAsFactors = FALSE)
  }
}


## Write output
for (i in 1:length(variables)) {
  ## First remove missing values and add variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]],
                                stringsAsFactors = FALSE)
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            cod = "George",
            nam = "George",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            repo = "",
            units = units[i],
            metaHead = ifelse(i==2, "PTC=?,PGC=F", ""),
            meta = Data[[variables[i]]][, 7],
            timef = 0)
}