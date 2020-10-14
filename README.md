# TerikoMacros
The *modules* directory contains VBA/macro code for performing useful tasks.
Each VBA/macro module is a file with a *.bas* filename extension.

To use a VBA/macro
1. Download the desired module
  * Ex: ManufacturerUtils.bas to your computer.
2. Open the Excel workbook intended to hold macros.
  * It's file extension is: *.xlsm*
3. Click *Developer* tab
4. Click the *Visual Basic* button on the Ribbon
  * This opens your Visual Basic editor.
5. Shift-Click on your project.
  * Ex: *VBAProject (TerikoMacros)*
6. On the drop-down box, click *Import File...*
7. Open or return to your working spreadsheet
  * NOT *TerikosMacros.xlsm*
8. Click *Developer* tab
9. Click *Macros* button
10. Find your desired macro on the list
  * Ex: *TerikoMacros!PrepareTable*
11. Click *Run* button

# ManufacturerUtils.bas
This module contains the following macros:
* PrepareTable

## ManufacturerUtils/PrepareTable
This macro assumes that the current sheet contains manufacturer data with a header row not already in *table* form.
Actions,
* Convert tabular data into an Excel *table* named *DataTable*
* Insert the following columns in before and based on the *Date* column
  * Qtr
    * Format: "Q<qtr>-<year>"
  * SY-Half
    * If <year> in Jan-Jun then Format: "2H-<year-1>-<year>"
    * If <year> in Jul-Dec then Format: "1H-<year>-<year+1>"
  * SY
    * If <year> in Jan-Jun then Format: "<year-1>-<year>"
    * If <year> in Jul-Dec then Format: "<year>-<year+1>"
  * Year
    * Format: <year>
* Insert the following columns based on *Manufacturer*, *#SKU*, *Cases (Product Detail)*
  * Item Description
    * insert before *PRODUCT_DESCRIPTION*
    * For each (Manufacturer, #SKU) pair, choose the *PRODUCT_DESCRIPTION* with the largest *Cases (Product Detail)* value
  * Item Pack
    * insert before *Pack Size*
    * For each (Manufacturer, #SKU) pair, choose the *Pack Size* with the largest *Cases (Product Detail)* value

That's all!
