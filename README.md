# SFDownloadFiles

SFDownloadFiles is an Excel Utility to download multiple files of all types from salesforce into a structured directory.

This utility was created to use a report of objects (Usually Cases) to download all files related to those objects into a structured directory with a folder for each object's files.

It works just as well if you just have a list of Content Document Links, or if you want, to download all files into the same directory.

For installation, configuration, and using this utility, see [SFDownloadFiles](https://github.com/jdpond/SFConversionsForExcel/wiki/SFDownloadFiles-User-Guide)

# SFconversionsExcelAddin

An add-in extension for Microsoft Excel containing conversion formulas:

1. **SFConvertId15to18** - convert a Salesforce 15 character ID to 18 character Case Safe Id. Equivalent of SF CASESAFEID())
1. **SFConvertId18to15** - convert a Salesforce 18 character Case Safe Id to 15 Character Id

## Install - Add as an Excel Extension

To enable this as an extension, download the [SFConversionsForExcelExtension.xlam](https://github.com/jdpond/SFUtilitiesForExcel/SFconversionsExcelAddin/blob/main/SFConversionsForExcel.xlam) file to your Excel Extension directory:

`C:\Users\[Your User]\AppData\Roaming\Microsoft\AddIns`

Then activate it with:

`Developer-->Excel Add-Ins-->SFconversionsExcelAddin(Checkbox)`

(You may have to enable the Developer tab on your ribbon by right clicking on the ribbon, Customize Ribbon-->Developer(Checkbox)

## Just the Code
The code is visible in the Visual basic file, [SFconversionsExcelAddin.bas]https://github.com/jdpond/SFUtilitiesForExcel/Sfconversionsforexcelextension/blob/main/SFConversionsForExcel.bas)
