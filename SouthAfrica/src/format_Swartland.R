# Converts the Swartland digitized data from the University of Witwatersrand 
# into the Station Exchange Format.
#
# Requires libraries XLConnect and dataresqc
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018
# Updated 7 Jan 2020

###############################################################################


require(XLConnect)
require(dataresqc)
options(scipen = 999) # avoid exponential notation


lat <- -33.4608
lon <- 18.7271
alt <- 123

inpath <- "../data/raw/Swartland/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("ta", "p", "dd")
units <- c("C", "hPa", "degree")
ids <- c(1086286, 1086285, 1086284)

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(convert_pressure(x, f = 25.4,
                                                   lat = lat, alt = alt), 1),
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
                          HH = integer(),
                          MM =integer(),
                          value = numeric(),
                          orig = character())
}


## Read data files
infiles <- list.files(inpath)
firstRow <- c(11, 10, 12, 13)
for (ifile in 1:length(infiles)) {
  year <- as.integer(substr(strsplit(infiles[ifile], "_")[[1]][6], 1, 4))
  
  ## Template 1 (1821-1822)
  if (year %in% 1821:1822) {
    template <- readWorksheetFromFile(paste0(inpath, infiles[ifile]), 
                                      startRow = firstRow[ifile],
                                      header = FALSE, sheet = 1, endCol = 6,
                                      colTypes = c("character",
                                                   "numeric",
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character"),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "p", "ta", "dd")
  
  ## Template 2 (1823-1824)
  } else if (year %in% 1823:1824) {
    template <- readWorksheetFromFile(paste0(inpath, infiles[ifile]), 
                                      startRow = firstRow[ifile],
                                      header = FALSE, sheet = 1, endCol = 5,
                                      colTypes = c("character",
                                                   "numeric",
                                                   "character",
                                                   "numeric",
                                                   "character"),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "ta", "dd")
    template$p <- NA
  }
  template$y <- year
  template$m <- as.integer(fill(get_month(template$m)))
  template$d <- fill(template$d)
  template <- template[which(!is.na(template$m)), ]

  ## Organize meta column
  template$p_orig <- paste0("Orig=", round(template$p, 2), "in")
  template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")
  template$dd_orig <- paste0("Orig=", template$dd)
  j <- which(!is.na(template$h))
  template[j, c("p_orig", "ta_orig", "dd_orig")] <-
    sapply(template[j, c("p_orig", "ta_orig", "dd_orig")],
           paste, paste0("orig.time=", template$h[j]), sep = "|")
  template$HH <- NA
  template$MM <- NA
  
  ## Transform wind direction to degrees
  template$dd <- sub(" ", "", template$dd)
  template$dd <- sub("to", "b", template$dd)
  template$dd <- sub("t", "b", template$dd)
  template$dd <- sub("by", "b", template$dd)
  directions <- c("N", "NhE", "NbE", "NbEhe", "NNE", "NNEhE", "NEbN", "NEhN", 
                  "NE", "NEhE", "NEbE", "NEbEhE", "ENE", "EbNhN", "EbN", "EhN",
                  "E", "EhS", "EbS", "EbShS", "ESE", "SEbEhE", "SEbE", "SEhE", 
                  "SE", "SEhS", "SEbS", "SSEhE", "SSE", "SbEhE", "SbE", "ShE",
                  "S", "ShW", "SbW", "SbWhW", "SSW", "SSWhW", "SWbS", "SWhS",
                  "SW", "SWhW", "SWbW", "SWbWhW", "WSW", "WbShS", "WbS", "WhS",
                  "W", "WhN", "WbN", "WbNhN", "WNW", "NWbWhW", "NWbW", "NWhW", 
                  "NW", "NWhN", "NWbN", "NNWhW", "NNW", "NbWhW", "NbW", "NhW")
  template$dd <- 5.625 * (match(template$dd, directions) - 1)

  ## Write to data frames
  for (i in 1:length(variables)) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c("y", "m", "d", "HH", "MM", variables[i], 
                                               paste0(variables[i], "_orig"))])
  }
}


## Write output
for (i in 1:length(variables)) {
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            cod = "Swartland",
            nam = "Swartland",
            variable = variables[i],
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            link = paste0("https://data-rescue.copernicus-climate.eu/lso/", ids[i]),
            units = units[i],
            stat = "point",
            metaHead = paste0("Data policy=GNU GPL v3.0", ifelse(i==2, "|PTC=N|PGC=Y", "")),
            meta = Data[[variables[i]]][, 7],
            period = 0)
}