# Rescued-Data
Data digitised from the capacity building regions of C3S 311a_Lot1 (aka [C3S Data Rescue Service](https://climate.copernicus.eu/data-rescue-service)).

### Structure
Each directory 'RegionName' has the following structure:

* data
  * raw: Digitisation templates as provided by the source
  * formatted: Data converted to SEF format
* docs: Documentation / Metadata
* src: Code used to format the data

### Data format
Standard filenames are used for formatted data:

```
<Source_Code>_<Station_Code>_<StartDate>_<EndDate>_<Variable>.tsv
```

Data are formatted into the [Station Exchange Format](https://github.com/C3S-Data-Rescue-Lot1-WP3/SEF/wiki) version 1.0.0, which implies one file for each variable.

### Variables
The following variables are provided:

* __ta__: air temperature (°C)
* __Tx__: maximum air temperature (°C)
* __Tn__: minimum air temperature (°C)
* __tb__: wet bulb temperature (°C)
* __td__: dew point (°C)
* __Ts__: soil temperature (°C)
* __p__: pressure (Pa)
* __mslp__: mean sea level pressure (Pa)
* __rh__: relative humidity (%)
* __rr__: precipitation (mm)
* __w__: wind speed (m/s)
* __wind_force__: wind force (Beaufort or other qualitative wind scales)
* __dd__: wind direction (°)
* __n__: cloud cover (%)
* __ss__: sunshine duration (hours)
* __vv__: horizontal visibility (1 to 9)

### Remarks
Note that the data have not been quality controlled.

### Contact
For any inquiry and to report bugs please write to [yuri.brugnara@giub.unibe.ch](mailto:yuri.brugnara@giub.unibe.ch).
