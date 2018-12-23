# Converts the Cape of Good Hope digitized data from the University of  
# Witwatersrand into the Station Exchange Format.
#
# Requires file write_sef.R and libraries XLConnect, plyr
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
require(plyr)
source("write_sef.R")
options(scipen = 999) # avoid exponential notation


lat <- ""
lon <- ""
alt <- ""

inpath <- "../data/raw/CapeOfGoodHope/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("ta", "p", "Tx", "Tn", "dd")
units <- c("C", "Pa", "C", "C", "degree")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(100 * convert_pressure(x, f = 25.4), 0),
                    Tx = function(x) round((x - 32) * 5 / 9, 1),
                    Tn = function(x) round((x - 32) * 5 / 9, 1),
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

# Define function to put each observation in a different row
split_obs <- function(x, v, h) { # x is the whole data frame, v is the variable to split,
                                 # h is a vector with the time labels
  j <- grep(v, names(x))
  n <- length(j)
  if (n != length(h)) stop("Wrong number of time labels")
  if (n <= 1) stop("There must be at least 2 columns for the same variable")
  x$h <- h[1]
  N <- dim(x)[1] / n
  for (i in 2:n) {
    x[((i-1)*N+1):(i*N), v] <- x[1:N, j[i]]
    x[((i-1)*N+1):(i*N), "h"] <- h[i]
  }
  x <- x[, -j[-1]]
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
infiles <- list.files(inpath)
for (ifile in 1:length(infiles)) {
  year <- as.integer(substr(rev(strsplit(infiles[ifile], "_")[[1]])[1], 1, 4))
  if (year %in% 1819:1821) {
    ## Format 1 (1819-1821)
    template <- readWorksheetFromFile(paste0(inpath, infiles[ifile]), 
                                      startRow = 11, header = FALSE, 
                                      sheet = 1, endCol = 8,
                                      colTypes = c("character",
                                                   rep("numeric", 5),
                                                   rep("character", 2)),
                                      drop = 3:4,
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "Tx", "Tn", "dd", "dd2")
    template <- rbind.fill(template, template[, 1:2]) # create new lines for the
                                                      # 2nd observation time
    template <- split_obs(template, "dd", c("AM", "PM"))
    template$Tx_orig <- paste0("Orig=", round(template$Tx, 0), "F")
    template$Tn_orig <- paste0("Orig=", round(template$Tn, 0), "F")
  } else if (year %in% 1822:1824) {
    ## Format 2 (1822-1824)
    template <- readWorksheetFromFile(paste0(inpath, infiles[ifile]), 
                                      startRow = 11, header = FALSE, 
                                      sheet = 1, endCol = 11,
                                      colTypes = c("character",
                                                   rep("numeric", 7),
                                                   rep("character", 3)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast") 
    names(template) <- c("m", "d", "p", "p2", "p3", "ta", "ta2", "ta3", 
                         "dd", "dd2", "dd3")
    template <- rbind.fill(template, template[, 1:2], template[, 1:2]) # 2nd + 3rd
                                                                       # obs. times
    template <- split_obs(template, "p", c("AM", "M", "PM"))
    template <- split_obs(template, "ta", c("AM", "M", "PM"))
    template <- split_obs(template, "dd", c("AM", "M", "PM"))
    template$p_orig <- paste0("Orig=", round(template$p, 1), "in")
    template$ta_orig <- paste0("Orig=", round(template$ta, 0), "F")
    template[ , c("p_orig", "ta_orig")] <- 
      sapply(template[ , c("p_orig", "ta_orig")],
             paste, paste0("t=", template$h), sep = ",")
  }
  template$dd_orig <- paste0("Orig=", template$dd)
  template$dd_orig <- paste0(template$dd_orig, ",t=", template$h)
  template$y <- year
  template$m <- as.integer(fill(get_month(template$m)))
  template$d <- fill(template$d)
  template <- template[which(!is.na(template$m)), ]
  template$h <- ""
  
  ## Transform wind direction to degrees
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(template$dd), directions) - 1)

  ## Write to data frames
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
  ## First remove missing values, order by time and add column with variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  Data[[variables[i]]] <- Data[[variables[i]]][order(Data[[variables[i]]]$y,
                                                     Data[[variables[i]]]$m,
                                                     Data[[variables[i]]]$d,
                                                     Data[[variables[i]]]$h), ]
  if (dim(Data[[variables[i]]])[1] > 0) {
    Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]],
                                  stringsAsFactors = FALSE)
    write_sef(Data = Data[[variables[i]]][, 1:6],
              outpath = outpath,
              cod = "Cape_Good_Hope",
              nam = "Cape of Good Hope",
              lat = lat,
              lon = lon,
              alt = alt,
              sou = "C3S_SouthAfrica",
              repo = "",
              units = units[i],
              metaHead = ifelse(i==2, "PTC=?,PGC=F", ""),
              meta = Data[[variables[i]]][, 7],
              timef = ifelse(variables[i] %in% c("Tx", "Tn"), "", 0))
  }
}