# Rescued-Data
Data digitized from the capacity building regions of C3S 311a_Lot1 (aka [C3S Data Rescue Service](https://climate.copernicus.eu/data-rescue-service)).

Each directory 'RegionName' has the following structure:

* data
  * raw: Digitisation templates as provided by the source
  * formatted: Data converted to SEF format
* docs: Documentation / Metadata
* src: Code used to format the data

Standard filenames are used for formatted data:

```
<Source_Code>_<Station_Code>_<StartDate>_<EndDate>_<Variable>.tsv
```

We use the [Station Exchange Format](http://brohan.org/SEF/SEF.html) version 0.0.1, which requires one file for each variable.

Note that the data provided here have not been quality controlled.
