# SFConversionsForExcel

An add-in extension for Microsoft Excel containing conversion formulas:

1. **SFConvertId15to18** - convert a Salesforce 15 character ID to 18 character Case Safe Id. Equivalent of SF CASESAFEID())
1. **SFConvertId18to15** - convert a Salesforce 18 character Case Safe Id to 15 Character Id

## Install - Add as an Excel Extension

To enable this as an extension, download the [SFConversionsForExcelExtension.xlam](https://github.com/jdpond/SFUtilitiesForExcel/Sfconversionsforexcelextension/blob/main/SFConversionsForExcel.xlam) file to your Excel Extension directory:

`C:\Users\[Your User]\AppData\Roaming\Microsoft\AddIns`

Then activate it with:

`Developer-->Excel Add-Ins-->Sfconversionsforcxcelextension(Checkbox)`

(You may have to enable the Developer tab on your ribbon by right clicking on the ribbon, Customize Ribbon-->Developer(Checkbox)

## Just the Code
The code is visible in the Visual basic file, [SFConversionsForExcelExtension.bas]https://github.com/jdpond/SFUtilitiesForExcel/Sfconversionsforexcelextension/blob/main/SFConversionsForExcel.bas)  