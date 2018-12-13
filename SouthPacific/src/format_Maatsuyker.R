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

variables <- c("t_air", "p")
units <- c("K", "Pa")
conversions <- list(t_air = function(x) round(273.15 + (x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(100 * convert_pressure(x, lat = lat, alt = alt), 0))


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
  names(template) <- c("y", "m", "d", "h", "p", "t_air")
  template <- template[which(!is.na(template$y)), ]
  template$p_orig <- paste0("Original=", round(template$p, 2), "in")
  template$t_air_orig <- paste0("Original=", round(template$t_air, 1), "F")
  
  ## Convert time to UTC
  hours <- as.integer(substr(template$h, 1, 2))
  hours <- hours + 10
  hours[hours > 23] <- hours[hours > 23] - 24
  template$d[hours <= 10] <- template$d[hours <= 10] + 1
  substr(template$h, 1, 2) <- formatC(hours, width = 2)
  
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
            cod = "C3S_Maatsuyker",
            nam = "Maatsuyker Lighthouse",
            lat = -43.656952,
            lon = 146.271518,
            alt = 120,
            sou = "C3S_SouthPacific",
            repo = "",
            units = units[i],
            metaHead = ifelse(i==2, "PTC=?", ""),
            meta = Data[[variables[i]]][, 7],
            timef = 0)
}