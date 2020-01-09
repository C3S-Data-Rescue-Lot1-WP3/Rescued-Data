# Converts the Maatsuyker Lighthouse digitized data from NIWA into the
# Station Exchange Format.
#
# Requires libraries XLConnect and dataresqc
#
# Created by Yuri Brugnara, University of Bern - 21 Dec 2018
# Updated 7 Jan 2020

###############################################################################


require(XLConnect)
require(dataresqc)
options(scipen = 999) # avoid exponential notation


lat <- -43.656952
lon <- 146.271518
alt <- 120

inpath <- "../data/raw/Maatsuyker/"
outpath <- "../data/formatted/"

variables <- c("ta", "p")
units <- c("C", "hPa")
ids <- c(1086325, 1086324)
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(convert_pressure(x, f = 25.4, lat = lat, alt = alt), 1))


## Initialize data frames
Data <- list()
for (v in variables) {
  Data[[v]] <- data.frame(year = integer(),
                          month = integer(),
                          day = integer(),
                          HH = integer(),
                          MM =integer(),
                          value = numeric(),
                          orig = character(),
                          time_orig = character())
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
  template$time_orig <- template$h
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
  template$HH <- as.integer(substr(template$h, 1, 2))
  template$MM <- as.integer(substr(template$h, 3, 4))
  
  ## Write to data frames
  for (i in 1:2) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                  template[, c(names(template)[1:3], "HH", "MM", variables[i], 
                                               paste0(variables[i], "_orig"), "time_orig")])
  }
}


## Write output
for (i in 1:2) {
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            variable = variables[i],
            cod = "Maatsuyker",
            nam = "Maatsuyker Lighthouse",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthPacific",
            link = paste0("https://data-rescue.copernicus-climate.eu/lso/", ids[i]),
            units = units[i],
            stat = "point",
            metaHead = paste0("Data policy=GNU GPL v3.0", ifelse(i==2, "|PTC=?|PGC=Y", "")),
            meta = paste0(Data[[variables[i]]][, 7], "|orig.time=", Data[[variables[i]]][, 8]),
            period = 0)
}