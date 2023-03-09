# Converts the Fort Napier (Pietermaritzburg) digitized data from the University of Witwatersrand 
# into the Station Exchange Format.
#
# Requires libraries XLConnect and dataresqc
#
# Created by Yuri Brugnara, University of Bern - 9 Mar 2023
# 

###############################################################################


require(XLConnect)
library(dataresqc)


lat <- -29.6128
lon <- 30.3696
alt <- 690

inpath <- "../data/raw/FortNapier/"
outpath <- "../data/formatted/"

# Define variables and units
variables <- c("ta", "p")
units <- c("C", "hPa")
ids <- c(NA, NA) # not in registry yet

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round((x - 32) * 5 / 9, 1),
                    p = function(x) 
                      round(convert_pressure(x, f = 25.4, 
                                                   lat = lat, alt = alt), 1))

## Initialize data frames
Data <- list()
for (v in variables) {
  Data[[v]] <- data.frame(year = integer(),
                          month = integer(),
                          day = integer(),
                          hour = numeric(),
                          minute = numeric(),
                          value = numeric(),
                          orig = character())
}


## Read data files
infile <- list.files(inpath, pattern="1875-1879", full.names=TRUE)
template <- readWorksheetFromFile(infile, startRow = 4, header = FALSE, sheet = 1, 
                                  startCol = 2, endCol = 10, readStrategy = "fast")
names(template) <- c("y", "m", "d", "c1", "c2", "p", "pmm", "ta", "tc")
template$HH <- 12
template$MM <- 43

## Organize meta column
template$p_orig <- paste0("Orig=", round(template$p, 2), "in")
template$ta_orig <- paste0("Orig=", round(template$ta, 1), "F")

## Write to data frames
Data <- list()
for (i in 1:length(variables)) {
  template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
  Data[[variables[i]]] <- rbind(Data[[variables[i]]], 
                                template[, c("y", "m", "d", "HH", "MM", variables[i], 
                                             paste0(variables[i], "_orig"))])
}


## Write output
for (i in 1:length(variables)) {
  write_sef(Data = Data[[variables[i]]][, 1:6],
            outpath = outpath,
            variable = variables[i],
            cod = "FortNapier",
            nam = "Pietermaritzburg (Fort Napier)",
            lat = lat,
            lon = lon,
            alt = alt,
            sou = "C3S_SouthAfrica",
#            link = paste0("https://data-rescue.copernicus-climate.eu/lso/", ids[i]),
            units = units[i],
            stat = "point",
            metaHead = paste0("Data policy=GNU GPL v3.0", ifelse(i==2, "|PTC=Y|PGC=Y", "")),
            meta = Data[[variables[i]]][, 7],
            period = 0,
            keep_na = TRUE)
}