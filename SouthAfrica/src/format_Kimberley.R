# Converts the Kimberley data digitized by
# the University of Witwatersrand into the Station Exchange Format.
#
# Requires file write_sef.R and libraries XLConnect, plyr
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
require(plyr)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


lat <- -28.715
lon <- 24.8375
alt <- 1234

inpath <- "../data/raw/Kimberley/"
outpath <- "../data/formatted/"

# Define variables, units, times
variables <- c("ta", "p", "tb", "Tx", "Tn", "dd", "wind_force", "rr", "td", "w")
units <- c("C", "Pa", "C", "C", "C", "degree", "", "mm", "C", "m/s")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x, y) 
                      round(100 * convert_pressure(x, f = 25.4, lat = lat, 
                                                   alt = alt), 0),
                    tb = function(x) round((x - 32) * 5 / 9, 1),
                    Tx = function(x) round((x - 32) * 5 / 9, 1),
                    Tn = function(x) round((x - 32) * 5 / 9, 1),
                    dd = function(x) round(x, 0),
                    wind_force = function(x) round(x, 0),
                    rr = function(x) round(x * 25.4, 1),
                    td = function(x) round((x - 32) * 5 / 9, 1),
                    w = function(x) round(x / 2.237, 1))

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

## Template 1 (Jan-Aug 1883)



for (infile in list.files(inpath, pattern = "xlsx")) {
  year <- as.integer(substr(strsplit(infile, "_")[[1]][2], 1, 4))
  
  ## Template 1 (1883)
  if (year == 1883) {
    ## Jan-Aug 1883
    template1 <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, endRow = 1894,
                                      header = FALSE,
                                      sheet = 1, endCol = 8,
                                      colTypes = c("character",
                                                   "numeric",
                                                   "character",
                                                   rep("numeric", 5)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template1) <- c("m", "d", "h", "ta", "tb", "Tn", "Tx", "p")
    ## Sep-Dec 1883
    template2 <- readWorksheetFromFile(paste0(inpath, infile), 
                                       startRow = 1897, header = FALSE,
                                       sheet = 1, endCol = 8,
                                       colTypes = c("character",
                                                    "numeric",
                                                    "character",
                                                    rep("numeric", 3),
                                                    rep("character", 2)),
                                       drop = 7,
                                       forceConversion = TRUE,
                                       readStrategy = "fast")
    names(template2) <- c("m", "d", "h", "ta", "tb", "p", "dd")
    template <- rbind.fill(template1, template2)

    ## Template 2 (1884-1885)  
  } else if (year %in% 1884:1885) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 1897 - year, header = FALSE,
                                      sheet = 1, endCol = 8,
                                      colTypes = c("character",
                                                   "numeric",
                                                   "character",
                                                   rep("numeric", 3),
                                                   rep("character", 2)),
                                      drop = 7,
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "ta", "tb", "p", "dd")
  
  ## Template 3 (1886-1889)  
} else if (year %in% 1886:1889) {
  template <- readWorksheetFromFile(paste0(inpath, infile), 
                                    startRow = 13, header = FALSE,
                                    sheet = 1, endCol = 15,
                                    colTypes = c("character",
                                                 "numeric",
                                                 "character",
                                                 rep("numeric", 7),
                                                 "character",
                                                 rep("numeric", 4)),
                                    drop = c(6, 7, 8, 10),
                                    forceConversion = TRUE,
                                    readStrategy = "fast")
  names(template) <- c("m", "d", "h", "ta", "tb", "p", "dd",
                        "wind_force", "rr", "Tx", "Tn")

  ## Template 4 (1898-1903)  
} else if (year %in% 1898:1903) {
  template <- readWorksheetFromFile(paste0(inpath, infile), 
                                    startRow = 11, header = FALSE,
                                    sheet = 1, endCol = 10,
                                    colTypes = c("character",
                                                 rep("numeric", 7),
                                                 "character",
                                                 "numeric"),
                                    forceConversion = TRUE,
                                    readStrategy = "fast")
  names(template) <- c("m", "d", "p", "Tx", "Tn", "ta", "tb", "td", "rr", "w")
  template$h <- ""
  template$rr <- as.numeric(sub("..", 0, template$rr, fixed = TRUE))
}
  
  template$m <- get_month(template$m)
  template$m[which(!template$m %in% as.character(1:12))] <- NA
  template$m <- as.integer(fill(template$m))
  template$d <- fill(template$d)
  template$h <- gsub(":", ".", template$h)
  template$h[grep("noon", template$h, ignore.case = TRUE)] <- "12.00"
  template$h[grep("midnight", template$h, ignore.case = TRUE)] <- "00.00"
  template$h[grep("8h25m26", template$h, ignore.case = TRUE)] <- "08.25"
  j <- grep("am", template$h, ignore.case = TRUE)
  template$h[j] <- format(as.numeric(sub(".{2}$", "", template$h[j])), nsmall = 2)
  j <- grep("pm", template$h, ignore.case = TRUE)
  template$h[j] <- format(as.numeric(sub(".{2}$", "", template$h[j])) + 12, nsmall = 2)
  template$p_orig <- paste0("Orig=", round(template$p, 3), "in")
  if (year < 1886) {
    template$p_orig <- paste(template$p_orig, "PTC=?", sep = ",")
  }
  template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")
  template$tb_orig <- paste0("Orig=", round(template$tb, 1), "F")
  if ("Tx" %in% names(template)) {
    template$Tx_orig <- paste0("Orig=", round(template$Tx, 1), "F")
    template$Tn_orig <- paste0("Orig=", round(template$Tn, 1), "F")
  }
  if ("wind_force" %in% names(template)) {
    template$wind_force_orig <- ""
  }
  if ("rr" %in% names(template)) {
    template$rr_orig <- paste0("Orig=", round(template$rr, 5), "in")
  }
  if ("td" %in% names(template)) {
    template$td_orig <- paste0("Orig=", round(template$td, 1), "F")
  }
  if ("w" %in% names(template)) {
    template$w_orig <- paste0("Orig=", round(template$w, 1), "mph")
  }

  
  ## Transform wind direction to degrees
  ## Entries like 'NWbN' are converted as 'NW'
  if ("dd" %in% names(template)) {
    template$dd_orig <- paste0("Orig=", template$dd)
    directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                    "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
    template$dd <- 22.5 * (match(toupper(sapply(strsplit(template$dd, "b"), 
                                                function(x) x[1])), directions) - 1)
  }

  
  ## Convert time to UTC (assuming local solar time is used)
  template$y <- year
  if (year < 1898) {
    dates <- paste(template$y, template$m, template$d, sep = "-")
    times <- strptime(paste(dates, template$h), 
                      format = "%Y-%m-%d %H.%M") - 3600 * 24 * lon / 360
    j <- which(template$h == "00.00")
    times[j] <- times[j] - 3600 * 24
    template$y <- as.integer(format(times, "%Y"))
    template$m <- as.integer(format(times, "%m"))
    template$d <- as.integer(format(times, "%d"))
    template$h <- format(times, "%H%M")
  }
  
  
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


## Order data by time and assign time flags
Tflags <- list()
for (v in c("ta", "p", "tb", "td", "dd", "wind_force", "w")) {
  Data[[v]] <- Data[[v]][order(Data[[v]]$y, Data[[v]]$m, Data[[v]]$d, Data[[v]]$h), ]
  Tflags[[v]] <- rep(0, dim(Data[[v]])[1])
}
for (v in c("Tx", "Tn")) {
  Data[[v]] <- Data[[v]][order(Data[[v]]$y, Data[[v]]$m, Data[[v]]$d, Data[[v]]$h), ]
  Tflags[[v]] <- rep(13, dim(Data[[v]])[1])
  Tflags[[v]][which(Data[[v]]$y %in% 1886:1889)] <- 6
  Data[[v]]$h[which(Data[[v]]$y %in% 1886:1889)] <- ""
  Tflags[[v]][which(Data[[v]]$y >= 1898)] <- 5
}
Tflags$rr <- rep(13, dim(Data$rr)[1])
Tflags$rr[which(Data$rr$y %in% 1886:1889)] <- 4
Data$rr$h[which(Data$rr$y %in% 1886:1889)] <- ""
Tflags$rr[which(Data$rr$y >= 1898)] <- 2
for (v in c("p", "ta", "tb", "td", "w")) {
  Tflags[[v]][which(Data[[v]]$y >= 1898)] <- 1
}


## Write output
for (i in 1:length(variables)) {
  ## First remove missing values and add column with variable code
  j <- which(!is.na(Data[[variables[i]]][, 5]))
  Data[[variables[i]]] <- Data[[variables[i]]][j, ]
  Tflags[[variables[i]]] <- Tflags[[variables[i]]][j]
  if (dim(Data[[variables[i]]])[1] > 0) {
    Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]],
                                  stringsAsFactors = FALSE)
    write_sef(Data = Data[[variables[i]]][, 1:6],
              outpath = outpath,
              cod = "Kimberley",
              nam = "Kimberley",
              lat = lat,
              lon = lon,
              alt = alt,
              sou = "C3S_SouthAfrica",
              repo = "",
              units = units[i],
              metaHead = ifelse(i==2, "PTC=T,PGC=T", ""),
              meta = Data[[variables[i]]][, 7],
              timef = Tflags[[variables[i]]])
  }
}