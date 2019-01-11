# Converts the Swellendam digitized data from the University of Witwatersrand 
# into the Station Exchange Format.
#
# Requires file write_sef.R and library XLConnect
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


lat <- -34.0257
lon <- 20.4381
alt <- 128

inpath <- "../data/raw/Swellendam/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("Tx", "Tn", "dd")
units <- c("C", "C", "degree")

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
                          time = character(),
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
  template$h <- ""
  template <- template[which(!is.na(template$m)), ]

  ## Organize meta column
  template$Tx_orig <- paste0("Orig=", round(template$Tx, 1), "F")
  template$Tn_orig <- paste0("Orig=", round(template$Tn, 1), "F")
  template$dd1_orig <- paste0("Orig=", template$dd1, ",t=AM")
  template$dd2_orig <- paste0("Orig=", template$dd2, ",t=PM")
  
  ## Transform wind direction to degrees
  ## Entries like 'NWbN' are converted as 'NW'
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd1 <- round(22.5 * (match(toupper(sapply(strsplit(template$dd1, "b"), 
                                               function(x) x[1])), directions) - 1), 0)
  template$dd2 <- round(22.5 * (match(toupper(sapply(strsplit(template$dd2, "b"), 
                                               function(x) x[1])), directions) - 1), 0)
  

  ## Write to data frames
  for (i in 1:length(variables)) {
    if (variables[i] == "dd") {
      template$dd <- template$dd1
      template$dd_orig <- template$dd1_orig
      Data$dd <- rbind(Data$dd, template[, c("y", "m", "d", "h", "dd", "dd_orig")])
      template$dd <- template$dd2
      template$dd_orig <- template$dd2_orig      
      Data$dd <- rbind(Data$dd, template[, c("y", "m", "d", "h", "dd", "dd_orig")])
    } else {
      template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
      Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                    template[, c("y", "m", "d", "h", variables[i], 
                                                 paste0(variables[i], "_orig"))])
    }
  }
  Data$dd <- Data$dd[order(Data$dd$y, Data$dd$m, Data$dd$d, Data$dd$h), ]
}


## Write output
for (i in 1:length(variables)) {
  ## First remove missing values and add variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]])
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            cod = "Swellendam",
            nam = "Swellendam",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            repo = "",
            units = units[i],
            metaHead = "",
            meta = Data[[variables[i]]][, 7],
            timef = ifelse(i == 3, 0, 13))
}