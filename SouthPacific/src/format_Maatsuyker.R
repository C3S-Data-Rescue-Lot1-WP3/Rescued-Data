# Converts the Maatsuyker Lighthouse digitized data from NIWA into the
# Station Exchange Format.
#
# Requires file write_sef.R and library XLConnect
#
# Created by Yuri Brugnara, University of Bern - 13 Dec 2018

###############################################################################


require(XLConnect)
source("write_sef.R")


lat <- -43.656952
lon <- 146.271518
alt <- 120

inpath <- "../data/raw/Maatsuyker/"
outpath <- "../data/formatted/"

variables <- c("ta", "p")
units <- c("C", "Pa")
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(100 * convert_pressure(x, f = 25.4, lat = lat, alt = alt), 0))


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
  template <- readWorksheetFromFile(paste0(inpath, infile), startRow = 4,
                                    header = FALSE, sheet = 1, endCol = 6,
                                    colTypes = c(rep("numeric", 3),
                                                 "character",
                                                 rep("numeric", 2)),
                                    forceConversion = TRUE)
  names(template) <- c("y", "m", "d", "h", "p", "ta")
  template <- template[which(!is.na(template$y)), ]
  template$h[template$h == "0000"] <- "2400"
  template$p_orig <- paste0("Orig=", round(template$p, 2), "in")
  template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")
  
  ## Convert time to UTC
  dates <- paste(template$y, template$m, template$d, sep = "-")
  times <- strptime(paste(dates, sub(" ", "0", template$h)), 
                    format = "%Y-%m-%d %H%M") - 3600 * 10
  template$y <- as.integer(format(times, "%Y"))
  template$m <- as.integer(format(times, "%m"))
  template$d <- as.integer(format(times, "%d"))
  template$h <- format(times, "%H%M")
  
  ## Write to data frames
  for (i in 1:2) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c(names(template)[1:4], variables[i], 
                                               paste0(variables[i], "_orig"))],
                                  stringsAsFactors = FALSE)
  }
}


## Write output
for (i in 1:2) {
  ## First remove missing values and add variable code
  Data[[variables[i]]] <- Data[[variables[i]]][which(!is.na(Data[[variables[i]]][, 5])), ]
  Data[[variables[i]]] <- cbind(variables[i], Data[[variables[i]]],
                                stringsAsFactors = FALSE)
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            cod = "Maatsuyker",
            nam = "Maatsuyker Lighthouse",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthPacific",
            repo = "",
            units = units[i],
            metaHead = ifelse(i==2, "PTC=?,PGC=T", ""),
            meta = Data[[variables[i]]][, 7],
            timef = 0)
}