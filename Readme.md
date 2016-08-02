# CREST Profile Extractor
This is a quick MATLAB script that uses a household load profile generation script, developed by Loughborough University, to populate an entire year worth of high-resolution data.

## The extractor runs on Windows!

For this code to work you will require MATLAB on Windows as well as Excel. The code interfaces with Excel through the `actxserver('Excel.Application')` to interface with Excel's COM and pass commands.

In order to run automatically, the original Excel Workbook has been downloaded from [here](https://dspace.lboro.ac.uk/dspace-jspui/handle/2134/5786) and the VBA Macros have been modified to no longer output any text box. Therefore, the MATLAB script can call all the required Excel functions in sequence without any user input requirements.

All code has been left untouched and the original `CREST_Domestic_electricity_demand_model_1.0e(1).xlsm` has been renamed to `modified.xlms`.

The output is a series (100 at the time of writing this) of MATLAB files that each contain the following information:

- Annual load profile in 1 minute resolution (kW)
- Annual solar irradiance profile in 1 minute resolution (kW / m^2)
- Number of residents in the household

## Sample image

![Sample profile for a random day](https://raw.githubusercontent.com/Muxelmann/CREST-profile-extractor/master/supporting/sample-profile.jpg)

## TODO

- Nothing really...


## DISCLAIMER

Despite me only having changed one line in VBA code (i.e. commented out a text box prompt), I do not take any responsibility if this does not work or even breaks!


