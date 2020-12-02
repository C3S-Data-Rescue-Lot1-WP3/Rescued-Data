# Converts the Cape Town Royal Observatory data digitized by
# the University of Witwatersrand into the Station Exchange Format.
#
# Requires libraries XLConnect, plyr, suncalc, dataresqc
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018
# Last updated 1 May 2020

###############################################################################


require(XLConnect)
require(plyr)
require(suncalc)
library(dataresqc)
options(scipen = 999) # avoid exponential notation


lat <- -33.9344
lon <- 18.4773
alt <- 12

inpath <- "../data/raw/CapeTownObservatory/"
outpath <- "../data/formatted/"

# Remove existing SEF files
file.remove(list.files(outpath, pattern="Cape_Town_Obs", full.names=TRUE))

# Define variables, units, times
variables <- c("ta", "p", "dd", "Tx", "Tn", "tb", "wind_force", "w", "n", 
               "rr", "ss", "vv", "w_max", "ta_mean", "ta2_mean", "ta3_mean",
               "tb_mean", "tb2_mean", "tb3_mean", "Tx2", "Tx3", "Tn2", "Tn3",
               "p_mean", "ta2", "ta3", "tb2", "tb3")
registry_id <- c(1086289, 1086288, 1086287, 1086292, 1086291, 1086293, 1086290, 
                 1086290, NA, NA, NA, NA, NA, 1086289, 1086289, 1086289,
                 1086293, 1086293, 1086293, 1086292, 1086292, 1086291, 1086291,
                 1086288, 1086289, 1086289, 1086293, 1086293)
units <- c("C", "hPa", "degree", "C", "C", "C", "Beaufort", "m/s", "%", 
           "mm", "hours", "1-9", "m/s", rep("C",10), "hPa", rep("C",4))
stats <- c(rep("point",3), "maximum", "minimum", rep("point",4), 
           rep("sum",2), "point", "maximum", rep("mean",6), rep("maximum",2),
           rep("minimum",2), "mean", rep("point",4))
periods <- c(rep(0,3), rep("p1day",2), rep(0,4), rep("day",2), 0, rep("day",12),
             rep(0,4))
types <- c("Sunrise", "Noon", "Sunset", "Midnight")
keeps <- c("sunrise", "solarNoon", "sunset", "nadir")

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    ta2 = function(x) round((x - 32) * 5 / 9, 1),
                    ta3 = function(x) round((x - 32) * 5 / 9, 1),
                    ta_mean = function(x) round((x - 32) * 5 / 9, 1),
                    ta2_mean = function(x) round((x - 32) * 5 / 9, 1),
                    ta3_mean = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x, y) 
                      round(convert_pressure(x, f = 25.4, lat = lat, 
                                             alt = alt, atb = y), 2),
                    p_mean = function(x) 
                      round(convert_pressure(x, f = 25.4, lat = lat, 
                                             alt = alt), 2),
                    dd = function(x) round(x, 0),
                    Tx = function(x) round((x - 32) * 5 / 9, 1),
                    Tx2 = function(x) round((x - 32) * 5 / 9, 1),
                    Tx3 = function(x) round((x - 32) * 5 / 9, 1),
                    Tn = function(x) round((x - 32) * 5 / 9, 1),
                    Tn2 = function(x) round((x - 32) * 5 / 9, 1),
                    Tn3 = function(x) round((x - 32) * 5 / 9, 1),
                    tb = function(x) round((x - 32) * 5 / 9, 1),
                    tb2 = function(x) round((x - 32) * 5 / 9, 1),
                    tb3 = function(x) round((x - 32) * 5 / 9, 1),
                    tb_mean = function(x) round((x - 32) * 5 / 9, 1),
                    tb2_mean = function(x) round((x - 32) * 5 / 9, 1),
                    tb3_mean = function(x) round((x - 32) * 5 / 9, 1),
                    wind_force = function(x) round(x, 0),
                    w = function(x) round(x / 2.237, 1),
                    w_max = function(x) round(x / 2.237, 1),
                    w_mean = function(x) round(x / 2.237, 1),
                    n = function(x) round(x * 20, 0),
                    rr = function(x) round(x * 25.4, 2),
                    ss = function(x) round(x, 2),
                    vv = function(x) round(x, 0))

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
for (year in 1834:1932) {
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
    template$h <- sub("1/4 to 1", "12:45", template$h)
    template$h <- sub("6pm", "18:00", template$h)
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
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", 
                              template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template[[paste0(dailyt, "_orig")]] <- paste0("Orig=", template[[dailyt]], "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    
    ## Template 3 (1857-1858)
    ## Structure change and new barometer in September 1858 
  } else if (year %in% 1857:1858) {
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
                         "ta", "tb", "Txn", "tb_new")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template <- template[which(!is.na(template$m)), ]
    if (year == 1858) {
      k <- which(template$m >= 9)
      template$ta[k] <- template$Txn[k]
      template$Txn[k] <- NA
      template$tb[k] <- template$tb_new[k]
    }
    fdiff <- template$Txn[2:dim(template)[1]] - template$Txn[1:(dim(template)[1]-1)]
    j <- which(fdiff > 0)
    template$Tn <- NA
    template$Tn[j] <- template$Txn[j]
    template$Tx <- NA
    template$Tx[j + 1] <- template$Txn[j + 1]
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", 
                              template$atb, "F")
    if (year == 1858) {
      template$p_orig[k] <- paste0(template$p_orig[k], "|instrument=Large Harvard Barometer")
    }
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
    template$p_orig <- paste0("Orig=", template$p, "in|instrument=Large Harvard Barometer")
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
    template$w <- c(rep(NA, 225), as.numeric(template$wind_force[226:dim(template)[1]]))
    template$wind_force[226:dim(template)[1]] <- NA
    template$p_orig <- paste0("Orig=", template$p, "in|instrument=Large Harvard Barometer")
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
    template$p_orig <- paste0("Orig=", template$p, "in|instrument=Large Harvard Barometer")
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
    template$p_orig <- paste0("Orig=", template$p, "in|instrument=Large Harvard Barometer")
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
    template$p_orig <- paste0("Orig=", template$p, "in|instrument=Large Harvard Barometer")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    
    ## Template 9 (1879) - two sheets (one for wind)
    ## Assuming the 2nd sheet is daily averages and pressure is already reduced to 32F
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
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", round(template$w, 2), "mph")
    template2 <- readWorksheetFromFile(paste0(inpath, infile), 
                                       startRow = 9, header = FALSE,
                                       sheet = 2,
                                       colTypes = c("character",
                                                    rep("numeric",14)),
                                       forceConversion = TRUE,
                                       readStrategy = "fast")
    names(template2) <- c("m", "d", "p_mean", 
                          "ta_mean", "tb_mean", "Tx_mean", "Tn_mean",
                          "ta2_mean", "tb2_mean", "Tx2", "Tn2",
                          "ta3_mean", "tb3_mean", "Tx3", "Tn3")
    template2$m <- as.integer(fill(get_month(template2$m)))
    template2$d <- fill(template2$d)
    template2 <- template2[which(!is.na(template2$m)), ]
    template2$p_mean_orig <- paste0("Orig=", template2$p_mean, "in")
    template2$ta_mean_orig <- paste0("Orig=", template2$ta_mean, "F|position=crib")
    template2$tb_mean_orig <- paste0("Orig=", template2$tb_mean, "F|position=crib")
    template2$Tx_orig <- paste0("Orig=", template2$Tx, "F|position=crib")
    template2$Tn_orig <- paste0("Orig=", template2$Tn, "F|position=crib")
    template2$ta2_mean_orig <- paste0("Orig=", template2$ta2_mean, "F|position=ground")
    template2$tb2_mean_orig <- paste0("Orig=", template2$tb2_mean, "F|position=ground")
    template2$Tx2_orig <- paste0("Orig=", template2$Tx2, "F|position=ground")
    template2$Tn2_orig <- paste0("Orig=", template2$Tn2, "F|position=ground")
    template2$ta3_mean_orig <- paste0("Orig=", template2$ta3_mean, "F|position=roof")
    template2$tb3_mean_orig <- paste0("Orig=", template2$tb3_mean, "F|position=roof")
    template2$Tx3_orig <- paste0("Orig=", template2$Tx3, "F|position=roof")
    template2$Tn3_orig <- paste0("Orig=", template2$Tn3, "F|position=roof")
    template2$h <- NA
    template <- rbind.fill(template, template2)
 
  ## Template 10 (1880) - Assuming these are daily averages  
  } else if (year == 1880) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                       startRow = 9, header = FALSE,
                                       sheet = 1, endCol = 7,
                                       colTypes = c("character",
                                                    rep("numeric",6)),
                                       forceConversion = TRUE,
                                       readStrategy = "fast")
    names(template) <- c("m", "d", "p_mean", "ta_mean", "tb_mean", "Tx", "Tn")
    template$m <- as.integer(fill(get_month(template$m)))
    template$d <- fill(template$d)
    template$h <- ""
    template <- template[which(!is.na(template$m)), ]
    template$p_mean_orig <- paste0("Orig=", template$p_mean, "in")
    template$ta_mean_orig <- paste0("Orig=", template$ta_mean, "F")
    template$tb_mean_orig <- paste0("Orig=", template$tb_mean, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    
    ## Template 11 (1881-1899)
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
                                      drop = 8,
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    names(template) <- c("m", "d", "h", "dd", "w", "p", "atb", "ta", "tb", "Tx",
                         "ta2", "tb2", "Tx2", "ta3", "tb3", "Tx3")
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
    template$Tn <- template$Tn2 <- template$Tn3 <- NA
    # Sort out Tn from Tx
    for (im in 1:12) {
      for (id in unique(template$d[which(template$m == im)])) {
        ih <- which(template$m == im & template$d == id)[1] + 
          which.min(template$Tx[which(template$m == im & template$d == id)]) - 1
        if (length(ih) > 0) {
          if (template$Tx[ih] <= min(template$ta[max(ih-2,1):ih],na.rm=TRUE)) {
            template$Tn[ih] <- template$Tx[ih]
            template$Tx[ih] <- NA
          }
        }
        ih <- which(template$m == im & template$d == id)[1] + 
          which.min(template$Tx2[which(template$m == im & template$d == id)]) - 1
        if (length(ih) > 0) {
          if (template$Tx2[ih] <= min(template$ta2[max(ih-2,1):ih],na.rm=TRUE)) {
            template$Tn2[ih] <- template$Tx2[ih]
            template$Tx2[ih] <- NA
          }
        }
        ih <- which(template$m == im & template$d == id)[1] + 
          which.min(template$Tx3[which(template$m == im & template$d == id)]) - 1
        if (length(ih) > 0) {
          if (template$Tx3[ih] <= min(template$ta3[max(ih-2,1):ih],na.rm=TRUE)) {
            template$Tn3[ih] <- template$Tx3[ih]
            template$Tx3[ih] <- NA
          }
        }
      }
    }
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F|screen=Window Crib")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F|screen=Window Crib")
    template$ta_orig <- paste0("Orig=", template$ta, "F|screen=Window Crib")
    template$tb_orig <- paste0("Orig=", template$tb, "F|screen=Window Crib")
    template$Tx2_orig <- paste0("Orig=", template$Tx2, "F|screen=Glaisher Stand")
    template$Tn2_orig <- paste0("Orig=", template$Tn2, "F|screen=Glaisher Stand")
    template$ta2_orig <- paste0("Orig=", template$ta2, "F|screen=Glaisher Stand")
    template$tb2_orig <- paste0("Orig=", template$tb2, "F|screen=Glaisher Stand")
    template$Tx3_orig <- paste0("Orig=", template$Tx3, "F|screen=Stevenson Crib")
    template$Tn3_orig <- paste0("Orig=", template$Tn3, "F|screen=Stevenson Crib")
    template$ta3_orig <- paste0("Orig=", template$ta3, "F|screen=Stevenson Crib")
    template$tb3_orig <- paste0("Orig=", template$tb3, "F|screen=Stevenson Crib")
    template$dd_orig <- paste0("Orig=", template$dd)
    if (year >= 1886) template$wind_force_orig <- paste0("Orig=", template$wind_force)
    else template$w_orig <- paste0("Orig=", template$w, "mph")
    
    ## Template 12 (1900-1902) - only 12-inch rain gauge read
  } else if (year %in% 1900:1902) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1,
                                      colTypes = c("character",
                                                   rep("numeric", 6),
                                                   rep("character", 2),
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 3),
                                                   rep("character", 2),
                                                   rep("numeric", 5),
                                                   rep("character", 2),
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 4)),
                                      keep = c(1:4,8:11,13,14,16:19,23:26,28,31),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    template[, 1] <- as.integer(fill(get_month(template[,1])))
    time1 <- c("0016", "0016", "0000")
    time2 <- c("0816", "0800", "0744")
    time3 <- c("2016", "2000", "1944")
    template1 <- template[, 1:8]
    names(template1) <- c("m", "d", "p", "atb", "dd", "wind_force", "n", "nh")
    template1$h <- time1[year-1899]
    template2 <- template[, c(1,2,9:12)]
    names(template2) <- c("m", "d", "p", "atb", "dd", "wind_force")
    template2$h <- time2[year-1899]
    template3 <- template[, c(1,2,13:20)]
    names(template3) <- c("m", "d", "p", "atb", "dd", "wind_force", "n", "nh", "rr", "ss")
    template3$h <- time3[year-1899]
    template <- rbind.fill(template1, template2, template3)
    template$n <- apply(template[,c("n","nh")], 1, max) # Take the max between high and low clouds
    if (year == 1900) template$rr[which(is.na(template$rr) & template$m<=6 & template$h==time3[1])] <- 0
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    template$n_orig <- paste0("Orig=", template$n)
    template$rr_orig <- paste0("Orig=", template$rr, "in")
    template$ss_orig <- paste0("Orig=", template$ss)
    
    ## Template 13 (1903-1923) - only 12-inch rain gauge read
    ## Wind run not formatted because of insufficient metadata
    ## From 1920 rain is given in 1/1000s of inch
  } else if (year %in% 1903:1923) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1,
                                      colTypes = c("character",
                                                   rep("numeric", 6),
                                                   rep("character", 2),
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 3),
                                                   rep("character", 2),
                                                   rep("numeric", 5),
                                                   rep("character", 2),
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 5)),
                                      drop = c(12,27),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    template <- template[, 1:26]
    template[, 1] <- as.integer(fill(get_month(template[,1])))
    time1 <- "0000"
    time2 <- "0744"
    time3 <- "1944"
    template1 <- template[, 1:11]
    names(template1) <- c("m", "d", "p", "atb", "ta", "tb", "Tn", "dd", "wind_force", "n", "nh")
    template1$h <- time1
    template2 <- template[, c(1,2,12:16)]
    names(template2) <- c("m", "d", "p", "atb", "ta", "dd", "wind_force")
    template2$h <- time2
    template3 <- template[, c(1,2,17:26)]
    names(template3) <- c("m", "d", "p", "atb", "ta", "tb", "Tx", "dd", "wind_force", "n", "nh", "rr")
    template3$h <- time3
    template <- rbind.fill(template1, template2, template3)
    template$n <- apply(template[,c("n","nh")], 1, max) # Take the max between high and low clouds
    template$rr[which(is.na(template$rr) & template$h==time3)] <- 0
    if (year >= 1920) template$rr <- template$rr / 1000
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    template$n_orig <- paste0("Orig=", template$n)
    template$rr_orig <- paste0("Orig=", template$rr, "in")
    
    ## Template 14 (1924) - only 12-inch rain gauge read
    ## Max wind speed and visibility begin in this year
  } else if (year == 1924) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1,
                                      colTypes = c("character",
                                                   rep("numeric", 6),
                                                   rep("character", 2),
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 3),
                                                   rep("character", 2),
                                                   rep("numeric", 5),
                                                   rep("character", 2),
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 5),
                                                   "character"),
                                      drop = c(12,27,29,30),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    template[, 1] <- as.integer(fill(get_month(template[,1])))
    time1 <- "0000"
    time2 <- "0744"
    time3 <- "1944"
    template1 <- template[, 1:11]
    names(template1) <- c("m", "d", "p", "atb", "ta", "tb", "Tn", "dd", "wind_force", "n", "nh")
    template1$h <- time1
    template2 <- template[, c(1,2,12:16)]
    names(template2) <- c("m", "d", "p", "atb", "ta", "dd", "wind_force")
    template2$h <- time2
    template3 <- template[, c(1,2,17:27)]
    names(template3) <- c("m", "d", "p", "atb", "ta", "tb", "Tx", "dd", "wind_force", "n", "nh", 
                          "vv", "rr")
    template3$h <- time3
    template4 <- template[, c(1,2,28,29)]
    names(template4) <- c("m", "d", "w_max", "dd_max")
    template4$h <- NA
    template <- rbind.fill(template1, template2, template3, template4)
    template$n <- apply(template[,c("n","nh")], 1, max) # Take the max between high and low clouds
    template$rr[which(is.na(template$rr) & template$h==time3)] <- 0
    template$rr <- template$rr / 1000
    template$w_max[which(template$m<5)] <- NA
    template$dd_max[which(template$m<5)] <- NA
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$wind_force_orig <- paste0("Orig=", template$wind_force)
    template$n_orig <- paste0("Orig=", template$n)
    template$vv_orig <- paste0("Orig=", template$vv)
    template$rr_orig <- paste0("Orig=", template$rr, "in")
    template$w_max_orig <- paste0("Orig=", template$w_max, "mph|dir=", template$dd_max)
    
    ## Template 15 (1925-1926) - only 12-inch rain gauge read
    ## Assumed rain is measured in the evening
    ## Time is given as S.A.S.T. (= UTC+2)
  } else if (year %in% 1925:1926) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = 1,
                                      colTypes = c("character",
                                                   rep("numeric", 6),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 6),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 4),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 4),
                                                   rep("character", 2)),
                                      drop = c(11,21,29,31,32),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    template[, 1] <- as.integer(fill(get_month(template[,1])))
    time1 <- "0830"
    time2 <- "1246"
    time3 <- "2030"
    template1 <- template[, 1:11]
    names(template1) <- c("m", "d", "p", "atb", "ta", "tb", "Tx", "dd", "w", "n", "vv")
    template1$h <- time1
    template2 <- template[, c(1,2,12:19)]
    names(template2) <- c("m", "d", "p", "atb", "ta", "tb", "Tn", "dd", "w", "n")
    template2$h <- time2
    template3 <- template[, c(1,2,20:27)]
    names(template3) <- c("m", "d", "p", "atb", "ta", "tb", "dd", "w", "n", "rr")
    template3$h <- time3
    template4 <- template[, c(1,2,28:30)]
    names(template4) <- c("m", "d", "w_max", "dd_max", "t_max")
    template4$h <- NA
    template <- rbind.fill(template1, template2, template3, template4)
    template$rr[which(is.na(template$rr) & template$h==time3)] <- 0
    template$rr <- template$rr / 1000
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    template$n_orig <- paste0("Orig=", template$n)
    template$vv_orig <- paste0("Orig=", template$vv)
    template$rr_orig <- paste0("Orig=", template$rr, "in")
    template$w_max_orig <- paste0("Orig=", template$w_max, "mph|dir=", template$dd_max, 
                                  "|time=", template$t_max)
    
    ## Template 16 (1927-1931) - only 12-inch rain gauge read
    ## Assumed rain is measured in the evening
    ## Time is given as S.A.S.T. (= UTC+2) until 1929, then as GMT
  } else if (year %in% 1927:1931) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 12, header = FALSE,
                                      sheet = ifelse(year<1930, 2, 1),
                                      colTypes = c("character",
                                                   rep("numeric", 6),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 5),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 4),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 4),
                                                   rep("character", 2)),
                                      drop = c(11,20,28,30,31),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    template[, 1] <- as.integer(fill(get_month(template[,1])))
    time1 <- "0830"
    time2 <- "1246"
    time3 <- "2030"
    template1 <- template[, 1:10]
    names(template1) <- c("m", "d", "p", "atb", "ta", "tb", "Tx", "dd", "w", "n")
    template1$h <- time1
    template2 <- template[, c(1,2,11:18)]
    names(template2) <- c("m", "d", "p", "atb", "ta", "tb", "Tn", "dd", "w", "n")
    template2$h <- time2
    template3 <- template[, c(1,2,19:26)]
    names(template3) <- c("m", "d", "p", "atb", "ta", "tb", "dd", "w", "n", "rr")
    template3$h <- time3
    template4 <- template[, c(1,2,27:29)]
    names(template4) <- c("m", "d", "w_max", "dd_max", "t_max")
    template4$h <- NA
    template <- rbind.fill(template1, template2, template3, template4)
    template$rr[which(is.na(template$rr) & template$h==time3)] <- 0
    if (year < 1930) template$rr <- template$rr / 1000
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    template$n_orig <- paste0("Orig=", template$n)
    template$rr_orig <- paste0("Orig=", template$rr, "in")
    template$w_max_orig <- paste0("Orig=", template$w_max, "mph|dir=", template$dd_max, 
                                  "|time=", template$t_max)
    
    ## Template 17 (1932) - only 12-inch rain gauge read
    ## Assumed rain is measured in the evening
    ## Time is given as C.M.T. (= UTC+1.5), probably a mistake
    ## There is a change of barometer on July 1st, atb units change to K
    ## There is an additional sheet with hourly wind, but it is not clear
    ## to which timezone they refer, so it was not formatted into SEF
  } else if (year == 1932) {
    template <- readWorksheetFromFile(paste0(inpath, infile), 
                                      startRow = 14, header = FALSE,
                                      sheet = 1,
                                      colTypes = c("character",
                                                   rep("numeric", 6),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 5),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 5),
                                                   "character",
                                                   rep("numeric", 2),
                                                   "character",
                                                   rep("numeric", 4),
                                                   rep("character", 2)),
                                      drop = c(11,20,25,29,31,32),
                                      forceConversion = TRUE,
                                      readStrategy = "fast")
    template[, 1] <- as.integer(fill(get_month(template[,1])))
    time1 <- "0830"
    time2 <- "1246"
    time3 <- "2030"
    template1 <- template[, 1:10]
    names(template1) <- c("m", "d", "p", "atb", "ta", "tb", "Tx", "dd", "w", "n")
    template1$h <- time1
    template2 <- template[, c(1,2,11:18)]
    names(template2) <- c("m", "d", "p", "atb", "ta", "tb", "Tn", "dd", "w", "n")
    template2$h <- time2
    template3 <- template[, c(1,2,19:26)]
    names(template3) <- c("m", "d", "p", "atb", "ta", "tb", "dd", "w", "n", "rr")
    template3$h <- time3
    template4 <- template[, c(1,2,27:29)]
    names(template4) <- c("m", "d", "w_max", "dd_max", "t_max")
    template4$h <- NA
    template <- rbind.fill(template1, template2, template3, template4)
    template$rr[which(is.na(template$rr) & template$h==time3 & template$m<=7)] <- 0
    template$p_orig <- paste0("Orig=", template$p, "in|atb=", template$atb, "F")
    template$ta_orig <- paste0("Orig=", template$ta, "F")
    template$tb_orig <- paste0("Orig=", template$tb, "F")
    template$Tn_orig <- paste0("Orig=", template$Tn, "F")
    template$Tx_orig <- paste0("Orig=", template$Tx, "F")
    template$dd_orig <- paste0("Orig=", template$dd)
    template$w_orig <- paste0("Orig=", template$w, "mph")
    template$n_orig <- paste0("Orig=", template$n)
    template$rr_orig <- paste0("Orig=", template$rr, "in")
    template$w_max_orig <- paste0("Orig=", template$w_max, "mph|dir=", template$dd_max, 
                                  "|time=", template$t_max)
    # July (Aug-Dec missing)
    j <- which(template[,1] == 7)
    template$p_orig[j] <- paste0("Orig=", template$p[j], "in|atb=", template$atb[j], 
                                 "K|instrument=Hicks 2131")
    template$atb[j] <- (template$atb[j] - 273.15) * 9 / 5 + 32


  }
  
  
  ## Transform wind direction to degrees
  if ("dd" %in% names(template)) {
    template$dd <- sub(" ", "", template$dd)
    template$dd <- sub("to", "b", template$dd)
    template$dd <- sub("t", "b", template$dd)
    template$dd <- sub("by", "b", template$dd)
    directions <- toupper(c("N", "NhE", "NbE", "NbEhe", "NNE", "NNEhE", "NEbN", "NEhN", 
                            "NE", "NEhE", "NEbE", "NEbEhE", "ENE", "EbNhN", "EbN", "EhN",
                            "E", "EhS", "EbS", "EbShS", "ESE", "SEbEhE", "SEbE", "SEhE", 
                            "SE", "SEhS", "SEbS", "SSEhE", "SSE", "SbEhE", "SbE", "ShE",
                            "S", "ShW", "SbW", "SbWhW", "SSW", "SSWhW", "SWbS", "SWhS",
                            "SW", "SWhW", "SWbW", "SWbWhW", "WSW", "WbShS", "WbS", "WhS",
                            "W", "WhN", "WbN", "WbNhN", "WNW", "NWbWhW", "NWbW", "NWhW", 
                            "NW", "NWhN", "NWbN", "NNWhW", "NNW", "NbWhW", "NbW", "NhW"))
    template$dd <- 5.625 * (match(toupper(template$dd), directions) - 1)
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
  
  
  ## Convert time to UTC (see file CapeTownObservatory_times.txt)
  template$y <- year
  template$h <- sub("h", "", template$h)
  if (year < 1858) {
    offset <- lon * 12 / 180
    tz <- ""
  } else if (year == 1858) {
    offset <- c(rep(lon*12/180,length(which(template$m<9))), 
                rep(0.662,length(which(template$m>=9))))
    tz <- c(rep("",length(which(template$m<9))), 
            rep("GoettingenMT",length(which(template$m>=9))))
  } else if (year == 1859) {
    offset <- 0.67
    tz <- "GoettingenMT"
  } else if (year %in% 1860:1880) {
    offset <- 1.5
    tz <- "CMT"
  } else if (year %in% 1881:1883) {
    offset <- 1.37 - 12 # all time labels are shifted by 12 hours
    tz <- ""
  } else if (year %in% 1884:1890) {
    template$h <- sub("1h", "13h", template$h) # 1PM was written as 1 instead of 13
    offset <- 1.37
    tz <- ""    
  } else if (year %in% 1891:1892) {
    offset <- 1.37
    tz <- ""
  } else if (year %in% 1893:1899) {
    offset <- 1.77
    tz <- ""
  } else if (year %in% 1900:1902) {
    offset <- 1.77 - 12 # all time labels are shifted by 12 hours
    tz <- ""
  } else if (year %in% 1903:1924) {
    offset <- 1.24 - 12 # all time labels are shifted by 12 hours
    tz <- ""
  } else if (year %in% 1925:1932) {
    offset <- 2
    tz <- "SAST"
  } 
  if (year != 1880) {
    template$h[which(nchar(template$h) == 3)] <- 
      paste0(0, template$h[which(nchar(template$h) == 3)])
    dates <- paste(template$y, template$m, template$d, sep = "-")
    cnames <- names(template)[grep("orig", names(template))]
    for (i in 1:length(cnames)) {
      ko <- grep("orig.time=", template[, cnames[i]])
      ko <- append(ko, which(is.na(template$h)))
      if (length(ko) > 0) {
        j <- (1:length(dates))[-ko]
      } else {
        j <- 1:length(dates)
      }
      if (length(tz) > 1) tz <- tz[j]
      template[j, cnames[i]] <- 
        paste0(template[j, cnames[i]], "|orig.time=", template$h[j], tz)
    }
    times <- strptime(paste(dates[j], template$h[j]), 
                      format = "%Y-%m-%d %H%M") - 3600 * offset
    template$y[j] <- as.integer(format(times, "%Y"))
    template$m[j] <- as.integer(format(times, "%m"))
    template$d[j] <- as.integer(format(times, "%d"))
    template$h[j] <- format(times, "%H%M")
    j <- grep(":", template$h)
    times <- strptime(paste(dates[j], template$h[j]), 
                      format = "%Y-%m-%d %H:%M") - 3600 * offset
    template$y[j] <- as.integer(format(times, "%Y"))
    template$m[j] <- as.integer(format(times, "%m"))
    template$d[j] <- as.integer(format(times, "%d"))
    template$h[j] <- format(times, "%H%M")
  }
  
  
  ## Convert units and write to data frames
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
      if (variables[i] == "n" & year >= 1925) {
        template[, variables[i]] <- template[, variables[i]] / 2
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
  if (variables[i]=="rr") { # From 1925 rain is measured at 9am
    Data[[variables[i]]]$hh[which(Data[[variables[i]]]$y>=1925)] <- 7
    Data[[variables[i]]]$mm[which(Data[[variables[i]]]$y>=1925)] <- 0
    Data[[variables[i]]][which(Data[[variables[i]]]$y>=1925), 5] <- 
      sub("0830", "0900", Data[[variables[i]]][which(Data[[variables[i]]]$y>=1925),5])
  }
  Data[[variables[i]]] <- Data[[variables[i]]][, c(1:3, 7:8, 5:6)]
  period <- periods[i]
  if (variables[i] %in% c("Tx","Tn")) {
    period <- rep(periods[i], nrow(Data[[variables[i]]]))
    period[which(Data[[variables[i]]][,1]%in%1903:1932)] <- "day"
  }
  vrb <- sub("_max", "", variables[i])
  vrb <- sub("_mean", "", vrb)
  vrb <- sub("2", "", vrb)
  vrb <- sub("3", "", vrb)
  note <- ""
  if (grepl("max",variables[i])) note <- "max"
  if (grepl("mean",variables[i])) note <- "mean"
  if (grepl("2",variables[i])) note <- paste0(note, "_bis")
  if (grepl("3",variables[i])) note <- paste0(note, "_ter")
  mHead <- "Data policy=GNU GPL v3.0"
  if (substr(variables[i],1,1) == "p") mHead <- paste0(mHead, "|PTC=Y|PGC=Y")
  if (variables[i] == "rr") mHead <- paste0(mHead, "|instrument=12-inch gauge")
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            variable = vrb,
            cod = "Cape_Town_Obs",
            nam = "Cape Town (Observatory)",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
            link = ifelse(is.na(registry_id[i]), NA,
                          paste0("https://data-rescue.copernicus-climate.eu/lso/", registry_id[i])),
            units = units[i],
            stat = stats[i],
            metaHead = ifelse(substr(variables[i],1,1)=="p", "PTC=Y|PGC=Y", ""),
            meta = Data[[variables[i]]][, 7],
            period = period,
            note = note)
}

## Rename SEF files
for (f in list.files(outpath, pattern="__", full.names=TRUE)) {
  file.rename(f, sub("__","_",f))
}

## Check SEF files
for (f in list.files(outpath, pattern="Cape_Town_Obs", full.names=TRUE)) {
  print(f)
  check_sef(f)
}