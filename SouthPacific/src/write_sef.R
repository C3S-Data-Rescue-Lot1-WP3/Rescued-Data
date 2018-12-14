#' Write data in Station Exchange Format
#'
#' @param Data A data frame with five or six variables (depending on time
#' resolution): variable code, year, month, day, (time in HHMM or HH:MM), value.
#' @param outpath Character string giving the output path (note that the filename
#' is generated from the source identifier, station code, start and end dates, 
#' and variable code).
#' @param cod Station code.
#' @param nam Station name.
#' @param lat Station latitude (degrees North in decimal).
#' @param lon Station longitude (degreees East in decimal).
#' @param alt Station altitude (metres).
#' @param sou Character string giving the source identifier.
#' @param repo Character string giving the repository identifier.
#' @param units Character string giving the units.
#' @param metaHead Character string giving any metadata for the header.
#' @param meta Character vector with length equal to the number of rows
#' of Data, giving any metadata for the single observations.
#' @param timef Integer giving the observation time period code. If NA 
#' (the default), it will be guessed from the dimension of Data and
#' the variable code.
#' @param note Character string to be added to the end of the filename.
#' It will be separated from the rest of the name by an underscore.
#' Blanks will be also replaced by underscores.
#' @param v Character string giving the SEF version.
#' 
#' @import utils
#' @export

write_sef <- function(Data, outpath, cod, nam, lat = NA, lon = NA, alt = NA, 
                      sou = NA, repo = NA, units = NA, metaHead = "", 
                      meta = "", timef = NA, note = "", v = "0.0.1") {
  
  ## Check which variables are given and produce one file per variable
  variables <- unique(Data[, 1])
  n <- length(variables)
  for (i in 1:n) {
    
    DataSubset <- subset(Data, Data[, 1] == variables[i], select = -1)
    
    ## Build filename
    datemin <- paste(formatC(unlist(Data[1, 2:4]), width=2, flag=0), 
                     collapse = "")
    datemax <- paste(formatC(unlist(Data[dim(Data)[1], 2:4]), width=2, flag=0), 
                     collapse = "")
    dates <- paste(datemin, datemax, sep = "-")
    filename <- paste(sou, cod, dates, variables[i], sep = "_")
    if (substr(outpath, nchar(outpath), nchar(outpath)) != "/") {
      outpath <- paste0(outpath, "/")
    }
    if (note != "") {
      gsub(" ", "_", note)
      note <- paste0("_", note)
    }
    filename <- paste0(outpath, filename, note, ".tsv")
  
    ## Build header
    header <- array(dim = c(11, 2), data = "")
    header[1, ] <- c("SEF", v)
    header[2, ] <- c("ID", cod)
    header[3, ] <- c("Name", nam)
    header[4, ] <- c("Lat", lat)
    header[5, ] <- c("Lon", lon)
    header[6, ] <- c("Alt", alt)
    header[7, ] <- c("Source", sou)
    header[8, ] <- c("Repo", repo)
    header[9, ] <- c("Var", variables[i])
    header[10, ] <- c("Units", units)
    header[11, ] <- c("Meta", metaHead)
    
    ## Guess time period of observation if not given
    if (is.na(timef)) {
      if (variables[i] == "rr") timef <- 12
      else if (dim(DataSubset)[2] == 4) timef <- 1
      else if (dim(DataSubset)[2] == 5) timef <- 0
    }
    
    ## Reshape data frame
    DataNew <- data.frame(Year = DataSubset[, 1],
                          Month = DataSubset[, 2],
                          Day = DataSubset[, 3],
                          HHMM = NA,
                          TimeF = timef,
                          Value = NA,
                          Meta = meta,
                          stringsAsFactors = FALSE)
    if (dim(DataSubset)[2] == 4) {
      DataNew$Value <- DataSubset[, 4]
    } else if (dim(DataSubset)[2] == 5) {
      DataNew$HHMM <- sub(":", "", DataSubset[, 4])
      DataNew$Value <- DataSubset[, 5]
    }
    
    ## Remove lines with missing data
    DataNew <- DataNew[which(!is.na(DataNew$Value)), ]
    
    ## Write header to file
    write.table(header, file = filename, quote = FALSE, row.names = FALSE, 
                col.names = FALSE, sep = "\t", dec = ".", fileEncoding = "UTF-8")
    
    ## Write column names to file
    write.table(t(names(DataNew)), file = filename, quote = FALSE, row.names = FALSE, 
                col.names = FALSE, sep = "\t", fileEncoding = "UTF-8",
                append = TRUE)
    
    ## Write data to file
    write.table(DataNew, file = filename, quote = FALSE, row.names = FALSE, 
                col.names = FALSE, sep = "\t", dec = ".", fileEncoding = "UTF-8", 
                append = TRUE)
  
  }
  
}


###############################################################################


#' Convert pressure data to hPa
#' 
#' Converts pressure observations made with a mercury barometer to SI units.
#' If geographical coordinates are given, a gravity correction is applied.
#'
#' @param p A vector of barometer observations in any unit of length.
#' @param f Conversion factor to mm (e.g., 25.4 for English inches).
#' @param lat Station latitude (degrees North in decimal).
#' @param alt Station altitude (metres).
#' 
#' @note
#' Ideally the barometer observations should be already reduced to 0 degrees Celsius.
#' 
#' @return
#' A vector of pressure values in hPa.
#' 
#' @references
#' WMO, 2008: Guide to meteorological instruments and methods of observation,
#' WMO-No. 8, World Meteorological Organization, Geneva.
#' 
#' @export

convert_pressure <- function(p, f = 25.4, lat = NA, alt = NA) {
  
  rho <- 13595.1
  lat <- lat * pi / 180
  
  g <- ifelse(!is.na(lat) & !is.na(alt),
              9.80620 * (1 - 0.0026442 * cos(2*lat) - 0.0000058 * (cos(2*lat))**2) 
                - 0.000003086 * alt,
              9.80665)
  
  p <- p * f * rho * g * 1e-05
  
  return(p)
  
}