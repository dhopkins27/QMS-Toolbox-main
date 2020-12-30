# QMS-toolbox

QMS toolbox contains python scripts to support large documenation projects and tracking of daily QMS metrics

## Installation

The .exe file can be run on any computer with admin access at IZI Medical.

The .py files needs to be editted on a computer with python installed

## Daily Tracker
This .exe and .py creates mailto links for QMS tracking metrics. 

Do not change path or location of ECO/NCMR/Deviation Log. And do not change which column data is stored in on the log

If you make changes to .py file, package it as .exe file using these instructions	https://datatofish.com/executable-pyinstaller/
```python
# change to correct file location in terminal
cd Desktop 
pyinstaller --onefile daily-tracker.py
# .exe is saved in the "dist" folder
```

## excelEditor

Script allows you to replace content in excel file reference template and save to numerous different files. Very useful to creating very similar routing files.

## wordEditor

Script allows you to replace content in a word file template and save to numerous different files.

## License
This toolbox was developed for IZI Medical internal use only
