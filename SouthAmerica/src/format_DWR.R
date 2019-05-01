# Converts the Daily Weather Reports digitized data for 1902 into the
# Station Exchange Format v0.2.
#
# Requires libraries readxl and dataresqc
# (https://github.com/c3s-data-rescue-service/dataresqc)
#
# Created by Yuri Brugnara, University of Bern - 30 Apr 2019

###############################################################################


require(readxl)
require(dataresqc)
options(scipen = 999) # avoid exponential notation


infile <- "../data/raw/DWR/South_America_1902.xlsx"
metafile <- "../data/raw/DWR/Positions.csv"
registryfile <- "../data/raw/DWR/c3s_argentina.csv"
outpath <- "../data/formatted/"

# Define variable codes and units
variables <- c("ta", "mslp", "Tx", "Tn", "rh", "dd", "n", "w")
units <- c("C", "Pa", "C", "C", "%", "degree", "%", "")
varids <- c(2, 0, 4, 5, 6, 7, 10, 8) # needed to build the link to the C3S registry

# Define conversions to apply to the raw data
conversions <- list(ta = function(x) round(x, 1),
                    mslp = function(x) round(100 * convert_pressure(x, f = 1), 0),
                    Tx = function(x) round(x, 1),
                    Tn = function(x) round(x, 1),
                    rh = function(x) round(x, 0),
                    dd = function(x) round(x, 0),
                    n = function(x) round(33.3 * x, 0),
                    w = function(x) round(x, 0))


## Read sheets names, station coordinates, registry entries
stations <- excel_sheets(infile)[-(1:2)]
coords <- read.csv(metafile)
registry <- read.csv(registryfile, sep = ";", encoding = "UTF-8")


## Loop on stations
for (ist in 1:length(stations)) {
  print(stations[ist])
  
  ## Initialize list of data frames (one per variable)
  Data <- list()
  
  ## Read data files
  template <- read_excel(infile, sheet = stations[ist], skip = 3,
                         col_types = c(rep("numeric", 3), rep("skip", 2),
                                       "numeric", "skip", "numeric", "skip",
                                       rep("numeric", 3), "text",
                                       "numeric", "skip", "numeric", "skip"))
  names(template) <- c("y", "m", "d", "mslp", "ta", "Tx", "Tn", 
                       "rh", "dd", "w", "n")
  template <- template[which(!is.na(template$y)), ]
  template$dd_orig <- paste0("Orig=", template$dd)
  template$n_orig <- paste0("Orig=", round(template$n, 0))
  template$mslp_orig <- paste0("Orig=", round(template$mslp, 1), "mm")
  template[, paste(c("ta", "Tx", "Tn", "rh", "w"), "orig", sep = "_")] <- ""
  
  ## Time is 2 PM LST until August and 7 AM LST from September
  template$HH <- NA
  template$HH[which(template$m <= 8)] <- 14
  template$HH[which(template$m > 8)] <- 7
  template$MM <- 0
  
  ## Transform wind direction to degrees
  directions <- c("N", "NNE", "NE", "ENE", "E", "ESE", "SE", "SSE", "S",
                  "SSW", "SW", "WSW", "W", "WNW", "NW", "NNW")
  template$dd <- 22.5 * (match(toupper(template$dd), directions) - 1)
  
  ## Convert units and write to data frame
  for (i in 1:length(variables)) {
    template[, variables[i]] <- conversions[[variables[i]]](template[, variables[i]])
    Data[[variables[i]]] <- as.data.frame(template[, c(names(template)[1:3], "HH", "MM", 
                                                       variables[i], 
                                                       paste0(variables[i], "_orig"))])
  }
  
  
  ## Output
  for (i in 1:length(variables)) {
    
    ## Define time statistic
    if (variables[i] == "Tx") {
      tstat <- "maximum"
    } else if (variables[i] == "Tn") {
      tstat <- "minimum"
    } else tstat <- "point"
    
    ## Add original time in meta column
    separator <- ifelse(Data[[variables[i]]][1,7] == "", "", "|")
    Data[[variables[i]]][which(Data[[variables[i]]]$m <= 8), 7] <- 
      paste(Data[[variables[i]]][which(Data[[variables[i]]]$m <= 8), 7],
            "orig.time=2pm", sep = separator)
    Data[[variables[i]]][which(Data[[variables[i]]]$m > 8), 7] <- 
      paste(Data[[variables[i]]][which(Data[[variables[i]]]$m > 8), 7],
            "orig.time=7am", sep = separator)
    
    ## Build ID for C3S register
    drsid <- registry[which(registry$Station.Name == stations[ist]), 1]
    drsid <- as.numeric(drsid) + varids[i] - 10
    
    ## Write file
    if (dim(Data[[variables[i]]])[1] > 0) {
      write_sef(Data = Data[[variables[i]]][, 1:6],
                outpath = outpath,
                variable = variables[i],
                cod = coords$SEF_ID[ist],
                nam = stations[ist],
                lat = coords$lat[ist],
                lon = coords$lon[ist],
                alt = coords$height[ist],
                sou = "C3S_SouthAmerica",
                link = paste0("https://data-rescue.copernicus-climate.eu/lso/", drsid),
                stat = tstat,
                units = units[i],
                metaHead = "Data policy=GNU GPL v3.0",
                meta = Data[[variables[i]]][, 7],
                period = ifelse(variables[i] %in% c("Tx","Tn"), "day", 0),
                time_offset = -4.29)
    }
  }
  
}