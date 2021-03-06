# Converts the Swellendam digitized data from the University of Witwatersrand 
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


lat <- -34.0257
lon <- 20.4381
alt <- 128

inpath <- "../data/raw/Swellendam/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("Tx", "Tn", "dd")
units <- c("C", "C", "degree")
ids <- c(1086320, 1086319, 1086318)
stats <- c("maximum", "minimum", "point")

# Define conversions to apply to the raw data
conversions <- list(Tx = function(x) round((x - 32) * 5 / 9, 1),
                    Tn = function(x) round((x - 32) * 5 / 9, 1))

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
                          MM = integer(),
                          value = numeric(),
                          orig = character())
}


## Read data files
for (year in 1821:1826) {
  infile <- paste0("Swellendam_", year, ".xlsx")
  template <- readWorksheetFromFile(paste0(inpath, infile), 
                                    startRow = 11, header = FALSE, 
                                    sheet = 1, endCol = 8,
                                    colTypes = c("character",
                                                 rep("numeric", 5),
                                                 rep("character",2)),
                                    drop = 3:4,
                                    forceConversion = TRUE,
                                    readStrategy = "fast")
  names(template) <- c("m", "d", "Tx", "Tn", "dd1", "dd2")
  template$y <- year
  template$m <- as.integer(fill(get_month(template$m)))
  template$d <- fill(template$d)
  template$HH <- NA
  template$MM <- NA
  template <- template[which(!is.na(template$m)), ]

  ## Organize meta column
  template$Tx_orig <- paste0("Orig=", round(template$Tx, 1), "F")
  template$Tn_orig <- paste0("Orig=", round(template$Tn, 1), "F")
  template$dd1_orig <- paste0("Orig=", template$dd1, "|orig.time=AM")
  template$dd2_orig <- paste0("Orig=", template$dd2, "|orig.time=PM")
  
  ## Transform wind direction to degrees
  directions <- c("N", "NhE", "NbE", "NbEhe", "NNE", "NNEhE", "NEbN", "NEhN", 
                  "NE", "NEhE", "NEbE", "NEbEhE", "ENE", "EbNhN", "EbN", "EhN",
                  "E", "EhS", "EbS", "EbShS", "ESE", "SEbEhE", "SEbE", "SEhE", 
                  "SE", "SEhS", "SEbS", "SSEhE", "SSE", "SbEhE", "SbE", "ShE",
                  "S", "ShW", "SbW", "SbWhW", "SSW", "SSWhW", "SWbS", "SWhS",
                  "SW", "SWhW", "SWbW", "SWbWhW", "WSW", "WbShS", "WbS", "WhS",
                  "W", "WhN", "WbN", "WbNhN", "WNW", "NWbWhW", "NWbW", "NWhW", 
                  "NW", "NWhN", "NWbN", "NNWhW", "NNW", "NbWhW", "NbW", "NhW")
  template$dd1 <- sub(" ", "", template$dd1)
  template$dd1 <- sub("to", "b", template$dd1)
  template$dd1 <- sub("t", "b", template$dd1)
  template$dd1 <- sub("by", "b", template$dd1)
  template$dd1 <- 5.625 * (match(template$dd1, directions) - 1)
  template$dd2 <- sub(" ", "", template$dd2)
  template$dd2 <- sub("to", "b", template$dd2)
  template$dd2 <- sub("t", "b", template$dd2)
  template$dd2 <- sub("by", "b", template$dd2)
  template$dd2 <- 5.625 * (match(template$dd2, directions) - 1)
  

  ## Write to data frames
  for (i in 1:length(variables)) {
    if (variables[i] == "dd") {
      template$dd <- template$dd1
      template$dd_orig <- template$dd1_orig
      Data$dd <- rbind(Data$dd, template[, c("y", "m", "d", "HH", "MM", "dd", "dd_orig")])
      template$dd <- template$dd2
      template$dd_orig <- template$dd2_orig      
      Data$dd <- rbind(Data$dd, template[, c("y", "m", "d", "HH", "MM", "dd", "dd_orig")])
    } else {
      template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
      Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                    template[, c("y", "m", "d", "HH", "MM", variables[i], 
                                                 paste0(variables[i], "_orig"))])
    }
  }
  Data$dd <- Data$dd[order(Data$dd$y, Data$dd$m, Data$dd$d), ]
}


## Write output
for (i in 1:length(variables)) {
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            variable = variables[i],
            cod = "Swellendam",
            nam = "Swellendam",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            link = paste0("https://data-rescue.copernicus-climate.eu/lso/", ids[i]),
            units = units[i],
            stat = stats[i],
            metaHead = "Data policy=GNU GPL v3.0",
            meta = Data[[variables[i]]][, 7],
            period = ifelse(i == 3, 0, "day"))
}