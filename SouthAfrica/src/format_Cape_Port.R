# Converts the Cape Town Port Office (Harbour Master) data digitized by
# the University of Witwatersrand into the Station Exchange Format.
#
# Years: 1834 to 1873, 1904
#
# Requires file write_sef.R and libraries XLConnect, plyr
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
require(plyr)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


lat <- -33.933
lon <- 18.477
alt <- 0

inpath <- "../data/raw/CapeTownPortOffice/"
outpath <- "../data/formatted/"

# Define variables, units, times
variables <- c("ta", "p", "dd", "wind_force")
units <- c("C", "Pa", "degree", "")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x, y) 
                      round(100 * convert_pressure(x, f = 25.4, lat = lat, 
                                                   alt = alt), 0),
                    dd = function(x) round(x, 0),
                    wind_force = function(x) round(x, 0))

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
  Data[[v]] <- data.frame(y = integer(),
                          m = integer(),
                          d = integer(),
                          h = character())
}


## Read data files
for (year in c(1829:1833, 1841:1850, 1855:1857, 1870:1873, 1904)) {
  infile <- paste0("Harbour_Master_", year, ".xlsx")
  
  ## Template 1 (1829-1873)
  if (year %in% 1829:1873) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 8, header = FALSE,
                                      sheet = 1, endCol = 6,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",2),
                                                   rep("numeric", 2)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "p", "ta")
    template$p_orig <- paste0("Orig=", round(template$p, 2), "in")
    template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")

    ## Template 2 (1904)  
  } else if (year == 1904) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 8, header = FALSE,
                                      sheet = 1, endCol = 5,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "wind_force")
    forces <- c("L", "F", "S", "V")
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    template$wind_force_orig[which(template$dd == "Calm")] <- "Orig=Calm"
    template$wind_force <- match(toupper(substr(template$wind_force, 1, 1)), forces)
    template$wind_force[which(template$dd == "Calm")] <- 0
    template$wind_force <- as.integer(template$wind_force)
  }
  
  template$m <- get_month(template$m)
  template$m[which(!template$m %in% as.character(1:12))] <- NA
  template$m <- as.integer(fill(template$m))
  template$d <- fill(template$d)
  template$dd_orig <- paste0("Orig=", template$dd)
  template$dd_orig[which(template$h == "AM")] <- 
    paste0("Orig=", template$dd[which(template$h == "AM")], ",t=AM")
  template$dd_orig[which(template$h == "PM")] <- 
    paste0("Orig=", template$dd[which(template$h == "PM")], ",t=PM")
  template$h[which(template$h %in% c("AM", "PM"))] <- ""
  template$h[grep("Noon", template$h, ignore.case = TRUE)] <- "1200"
  j <- grep("AM", template$h, ignore.case = TRUE)
  template$h[j] <- paste0("0", substr(template$h[j], 1, 1), "00")
  j <- grep("PM", template$h, ignore.case = TRUE)
  template$h[j] <- paste0(as.integer(substr(template$h[j], 1, 1)) + 12, "00")

  
  ## Transform wind direction to degrees
  ## Entries like 'NWbN' are converted as 'NW'
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(sapply(strsplit(template$dd, "b"), 
                                              function(x) x[1])), directions) - 1)

  
  ## Convert time to UTC (assuming local solar time is used)
  template$y <- year
  dates <- paste(template$y, template$m, template$d, sep = "-")
  j <- which(nchar(template$h) == 4)
  times <- strptime(paste(dates[j], template$h[j]), 
                    format = "%Y-%m-%d %H%M") - 3600 * 24 * lon / 360
  template$y[j] <- as.integer(format(times, "%Y"))
  template$m[j] <- as.integer(format(times, "%m"))
  template$d[j] <- as.integer(format(times, "%d"))
  template$h[j] <- format(times, "%H%M")
  
  
  ## Write to data frames
  template <- template[which(!is.na(template$m)), ]
  for (i in 1:length(variables)) {
    if (variables[i] %in% names(template)) {
      template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
      Data[[variables[i]]] <- rbind.fill(Data[[variables[i]]], 
                                         template[, c("y", "m", "d", "h", variables[i], 
                                                      paste0(variables[i], "_orig"))])
    }
  }
  
}


## Write output
for (i in 1:length(variables)) {
  ## First remove missing values and add column with variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  if (dim(Data[[variables[i]]])[1] > 0) {
    Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]],
                                  stringsAsFactors = FALSE)
    write_sef(Data = Data[[variables[i]]][, 1:6],
              outpath = outpath,
              cod = "Cape_Town_Port",
              nam = "Cape Town (Port Office)",
              lat = lat,
              lon = lon,
              alt = alt,
              sou = "C3S_SouthAfrica",
              repo = "",
              units = units[i],
              metaHead = ifelse(i==2, "PTC=?,PGC=T", ""),
              meta = Data[[variables[i]]][, 7],
              timef = 0)
  }
}