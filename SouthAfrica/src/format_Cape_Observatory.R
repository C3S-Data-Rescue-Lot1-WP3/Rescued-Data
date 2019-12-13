# Converts the Cape Town Royal Observatory data digitized by
# the University of Witwatersrand into the Station Exchange Format.
#
# Requires libraries XLConnect, plyr, suncalc, dataresqc
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018

###############################################################################


require(XLConnect)
require(plyr)
require(suncalc)
library(SEF)
#library(dataresqc)
options(scipen = 999) # avoid exponential notation


lat <- -33.9344
lon <- 18.4773
alt <- 12

inpath <- "../data/raw/CapeTownObservatory/"
outpath <- "../data/formatted/"

# Define variables, units, times
variables <- c("ta", "p", "dd", "Tx", "Tn", "tb", "wind_force", "w")
registry_id <- c(1086289, 1086288, 1086287, 1086292, 1086291, 1086293, 1086290, 1086290)
units <- c("C", "hPa", "degree", "C", "C", "C", "Beaufort", "m/s")
stats <- c(rep("point",3), "maximum", "minimum", rep("point",2), "mean")
periods <- c(rep(0,3), rep("p1day",2), rep(0,2), "p1day")
types <- c("Sunrise", "Noon", "Sunset", "Midnight")
keeps <- c("sunrise", "solarNoon", "sunset", "nadir")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x, y) 
                      round(dataresqc::convert_pressure(x, f = 25.4, lat = lat, 
                                                   alt = alt, atb = y), 2),
                    dd = function(x) round(x, 0),
                    Tx = function(x) round((x - 32) * 5 / 9, 1),
                    Tn = function(x) round((x - 32) * 5 / 9, 1),
                    tb = function(x) round((x - 32) * 5 / 9, 1),
                    wind_force = function(x) round(x, 0),
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

# Define function to replace descriptive times (such as 'sunrise') with UTC times
to_utc <- function(x, keep, lat, lon, y) { # x is a data frame with columns m,d,h
  dates <- as.Date(paste(y, x$m, x$d, sep = "-"), format = "%Y-%m-%d")
  times <- getSunlightTimes(dates, lat, lon, keep = keep)
  x$h <- format(times[[keep]], format = "%Hh%M")
  return(x$h)
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
for (year in 1834:1899) {
  infile <- paste(year, "xlsx", sep = ".")
  
  ## Template 1 (1834-1842)
  if (year %in% 1834:1842) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1, endCol = 11,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3),
                                                   rep("numeric", 6)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "wind_force", "p", "atb", 
                         "ta", "tb", "Tx", "Tn")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", 
                              template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
#    j <- which(template$h %in% types)
#    if (length(j) > 0) {
      template[, paste0(c("p", "ta", "tb", "Tx", "Tn", "dd", "wind_force"), "_orig")] <-
        sapply(template[, paste0(c("p", "ta", "tb", "Tx", "Tn", "dd", "wind_force"), 
                                  "_orig")],
               paste, paste0("orig.time=", template$h), sep = "|")
      for (i in 1:4) {
        j <- which(template$h == types[i])
        if (length(j) > 0) {
          template$h[j] <- to_utc(template[j, ], keeps[i], lat, lon, year)
        }
      }
#    }
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    
    ## Template 2 (1843-1856)  
  } else if (year %in% 1843:1856) {
    if (year <= 1847 | year %in% 1850:1852) firstRow <- 11
    if (year %in% 1848:1849) firstRow <- 10
    if (year >= 1853) firstRow <- 12
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = firstRow, header = FALSE,
                                      sheet = 1, endCol = 10,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3),
                                                   rep("numeric", 5)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    dailyt <- ifelse(year == 1856, "Tx", "Tn")
    names(template) <- c("m", "d", "h", "dd", "wind_force", "p", "atb", 
                         "ta", "tb", dailyt)
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", 
                              template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template[[paste0(dailyt, "_orig")]] <- paste0("Orig=", template[[dailyt]], "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    
    ## Template 3 (1857-1858)  
  } else if (year %in% 1857:1858) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1, endCol = 10,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3),
                                                   rep("numeric", 5)),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "wind_force", "p", "atb", 
                         "ta", "tb", "Txn")
    fdiff <- template$Txn[2:dim(template)[1]] - template$Txn[1:(dim(template)[1]-1)]
    j <- which(fdiff > 0)
    template$Tn <- NA
    template$Tn[j] <- template$Txn[j]
    template$Tx <- NA
    template$Tx[j + 1] <- template$Txn[j + 1]
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", 
                              template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    
    ## Template 4 (1859)  
  } else if (year == 1859) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1, endCol = 12,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3),
                                                   rep("numeric", 7)),
                                      drop = c(6, 7, 9, 10),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "wind_force", "p", "ta", "tb")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$p_orig <- paste0("Orig=", template$p, "in")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    
    ## Template 5 (1860)  
  } else if (year == 1860) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1, endCol = 11,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3),
                                                   rep("numeric", 6)),
                                      drop = c(6, 7, 9),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "wind_force", "p", "ta", "tb")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$w <- c(rep(NA, 225), as.numeric(template$wind_force[226:dim(template)[1]]))
    template$wind_force[226:dim(template)[1]] <- NA
    template$p_orig <- paste0("Orig=", template$p, "in")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    
    ## Template 6 (1861-1873)  
  } else if (year %in% 1861:1873) {
    if (year <= 1868) firstRow <- 12
    if (year == 1870) firstRow <- 10
    if (year %in% c(1869, 1871:1873)) firstRow <- 11
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = firstRow, header = FALSE,
                                      sheet = 1, endCol = 11,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",2),
                                                   rep("numeric", 7)),
                                      drop = c(6, 7, 9),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "w", "p", "ta", "tb")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$p_orig <- paste0("Orig=", template$p, "in")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    
    ## Template 7 (1874-1876)  
  } else if (year %in% 1874:1876) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1, endCol = 13,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",3),
                                                   rep("numeric", 8)),
                                      drop = c(6, 7, 11),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "wind_force", "p", "Tx", "Tn", "ta", "tb")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$p_orig <- paste0("Orig=", template$p, "in")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    
    ## Template 8 (1877-1878)  
  } else if (year %in% 1877:1878) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 11, header = FALSE,
                                      sheet = 1, endCol = 13,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",2),
                                                   rep("numeric", 9)),
                                      drop = c(6, 7, 11),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "w", "p", "Tx", "Tn", "ta", "tb")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$p_orig <- paste0("Orig=", template$p, "in")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    
    ## Template 9 (1879) - two sheets - only "ground" thermometers read
  } else if (year == 1879) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1, endCol = 5,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",2),
                                                   "numeric"),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "w")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", round(template$w, 2), "mph")
    template2 <- readWorksheetFromFile(paste0(inpath, infile), 
                                       startRow = 9, header = FALSE,
                                       sheet = 2, endCol = 11,
                                       colTypes = c("character",
                                                    rep("numeric",10)),
                                       drop = 4:7,
                                       forceConversion = TRUE,
                                       readStrategy = "fast")
    names(template2) <- c("m", "d", "p", "ta", "tb", "Tx", "Tn")
    template2$m <- as.integer(fill(get_month(template2$m)))
    template2$d <- fill(template2$d)
    template2 <- template2[which(!is.na(template2$m)), ]
    template2$p_orig <- paste0("Orig=", template2$p, "in|PTC=?")
    template2$ta_orig <- paste0("Orig=", template2$ta, "F")
    template2$tb_orig <- paste0("Orig=", template2$tb, "F")
    template2$Tx_orig <- paste0("Orig=", template2$Tx, "F")
    template2$Tn_orig <- paste0("Orig=", template2$Tn, "F")
    template2$h <- "130"
    template <- merge(template, template2, by.x = c("m", "d", "h"), all.x = TRUE)
 
  ## Template 10 (1880) - only "ground" thermometers read     
  } else if (year == 1880) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                       startRow = 9, header = FALSE,
                                       sheet = 1, endCol = 7,
                                       colTypes = c("character",
                                                    rep("numeric",6)),
                                       forceConversion = TRUE,
                                       readStrategy = "fast")
    names(template) <- c("m", "d", "p", "ta", "tb", "Tx", "Tn")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template$h <- ""
    template <- template[which(!is.na(template$m)), ]
    template$p_orig <- paste0("Orig=", template$p, "in|PTC=?")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    
    ## Template 11 (1881-1899) - only "Stevenson Crib" temperature read 
  } else if (year %in% 1881:1899) {
    if (year <= 1884 | year == 1890) firstRow <- 11
    if (year == 1885 | year >= 1893) firstRow <- 9
    if (year %in% c(1886, 1888, 1891, 1892)) firstRow <- 12
    if (year %in% c(1887, 1889)) firstRow <- 13
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = firstRow, header = FALSE,
                                      sheet = 1, endCol = 17,
                                      colTypes = c("character",
                                                   "numeric",
                                                   rep("character",2),
                                                   rep("numeric", 13)),
                                      drop = 8:14,
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "w", "p", "atb", "ta", "tb", "Tx")
    if (year >= 1886) names(template)[5] <- "wind_force"
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    template$h <- sub("h", "", template$h)
    template$h <- sub(":", "", template$h)
    if (year >= 1893) {
      substr(template$h[grep("pm", template$h, ignore.case = TRUE)], 1, 2) <- 
        as.character(as.integer(substr(template$h[grep("pm", template$h, ignore.case = TRUE)], 
                                       1, 2)) + 12)
      template$h <- substr(template$h, 1, 4)
    }
    template$Tn <- NA
    for (im in 1:12) {
      for (id in unique(template$d[which(template$m == im)])) {
        if (year != 1881 & im != 1) {
          ih <- which(template$m == im & template$d == id)[1] + 
            which.min(template$Tx[which(template$m == im & template$d == id)]) - 1
          template$Tn[ih] <- template$Tx[ih]
          template$Tx[ih] <- NA
        }
      }
    }
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    if (year >= 1886) template$wind_force_orig <- paste0("Orig=", template$wind_force)
    else template$w_orig <- paste0("Orig=", template$w, "mph")
    
  }
  
  
  ## Transform wind direction to degrees
  ## Entries like 'NW to N' are converted as 'NW', entries like 'NWhN' as NA
  if ("dd" %in% names(template)) {
    directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                    "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
    template$dd <- 22.5 * (match(toupper(sapply(strsplit(template$dd, " "), 
                                                function(x) x[1])), directions) - 1)
  }
  
  ## Take the average of wind force for entries like '6 to 8'
  if ("wind_force" %in% names(template)) {
    wf <- array(dim = c(dim(template)[1], 2))
    wf[, 1] <- as.integer(sapply(strsplit(as.character(template$wind_force), " "), 
                                 function(x) x[1]))
    wf[, 2] <- as.integer(sapply(strsplit(as.character(template$wind_force), " "), 
                                 function(x) x[3]))
    if (sum(!is.na(wf[, 2])) > 0) {
      template$wind_force[which(!is.na(rowMeans(wf)))] <- 
        rowMeans(wf)[which(!is.na(rowMeans(wf)))]
    }
    template$wind_force <- as.numeric(template$wind_force)
  }
  
  
  ## Convert time to UTC (assuming local solar time is used)
  template$y <- year
  if (year != 1880) {
    template$h[which(nchar(template$h) == 3)] <- 
      paste0(0, template$h[which(nchar(template$h) == 3)])
    dates <- paste(template$y, template$m, template$d, sep = "-")
    cnames <- names(template)[grep("orig", names(template))]
    for (i in 1:length(cnames)) {
      ko <- grep("orig.time=", template[, cnames[i]])
      if (length(ko) > 0) {
        j <- (1:length(dates))[-ko]
      } else {
        j <- 1:length(dates)
      }
      template[j, cnames[i]] <- 
        paste0(template[j, cnames[i]], "|orig.time=", template$h[j])
    }
    times <- strptime(paste(dates[j], template$h[j]), 
                      format = "%Y-%m-%d %H%M") - 3600 * 24 * lon / 360
    template$y[j] <- as.integer(format(times, "%Y"))
    template$m[j] <- as.integer(format(times, "%m"))
    template$d[j] <- as.integer(format(times, "%d"))
    template$h[j] <- format(times, "%H%M")
  }
  
  
  ## Write to data frames
  template <- template[which(!is.na(template$m)), ]
  for (i in 1:length(variables)) {
    if (variables[i] %in% names(template)) {
      if (variables[i] == "p") {
        if (year <= 1858 | year >= 1881) {
          template[, variables[i]] <- 
            conversions[[variables[i]]](template[, variables[i]], template$atb)
        } else {
          template[, variables[i]] <- 
            conversions[[variables[i]]](template[, variables[i]], NULL)
        }
      } else {
        template[, variables[i]] <- 
          conversions[[variables[i]]](template[, variables[i]])
      }
      Data[[variables[i]]] <- rbind.fill(Data[[variables[i]]], 
                                         template[, c("y", "m", "d", "h", variables[i], 
                                                      paste0(variables[i], "_orig"))])
      if (!variables[i] %in% c("dd", "w") & year == 1879) {
        Data[[variables[i]]]$h[which(Data[[variables[i]]]$y == year)] <- ""
      }
    }
  }
  
}


## Write output
for (i in 1:length(variables)) {
  ## Order by time and split hour and minute
  Data[[variables[i]]] <- Data[[variables[i]]][order(Data[[variables[i]]]$y,
                                                     Data[[variables[i]]]$m,
                                                     Data[[variables[i]]]$d,
                                                     Data[[variables[i]]]$h), ]
  Data[[variables[i]]]$hh <- as.integer(substr(Data[[variables[i]]]$h, 1, 2))
  Data[[variables[i]]]$mm <- as.integer(substr(Data[[variables[i]]]$h, 3, 4))
  Data[[variables[i]]] <- Data[[variables[i]]][, c(1:3, 7:8, 5:6)]
  ## Remove data for 1879-1880 because we don't know what they are
  if (!variables[i] %in% c("Tx","Tn")) {
    Data[[variables[i]]] <- Data[[variables[i]]][which(!Data[[variables[i]]]$y %in% 1879:1880), ]
  }
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            variable = variables[i],
            cod = "Cape_Town_Obs",
            nam = "Cape Town (Observatory)",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            link = paste0("https://data-rescue.copernicus-climate.eu/lso/", registry_id[i]),
            units = units[i],
            stat = stats[i],
            metaHead = paste0("Data policy=GNU GPL v3.0", ifelse(i==2, "|PTC=Y|PGC=Y", "")),
            meta = Data[[variables[i]]][, 7],
            period = periods[i])
}