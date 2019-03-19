# Excel Module

## Description
This module serve for createing Excel Files. It primarily developed on top of ExcelJS Node Module.
It provides highlevel methods to easly create excel files.

## Configure Project
-  Clone this project.
-   Perform 
    ```
    yarn install
    ```
-   Perform
    ```
    yarn link
    ```
-   Now the project is configured to be used as a local module.

## How to Use ?

```javascript

const excel = require('excelmodule')

// Create Workbook 
let xlWorkbook = new excel.ExcelWorkbook();

// Create Worksheet 
let xlWorksheet = new excel.ExcelWorksheet(xlWorkbook,'SheetName',{
    properties : {
        tabColor : {
            argb : 'FFFF0000'
        },
        showGridLines : false
    }
}) 

// Set Columns Headers to table
xlWorksheet.addHeaders([
    {header : 'S.NO.', key : 'sno', width: 20},
    {header : 'Items', key : 'items', width: 40},
    {header : 'Quantity', key : 'qty', width: 20},
])

// Set Column Header Color
xlWorksheet.colorHeader({
    argb : 'FFFF0000'
})

// Define data 
let data = [
    {
        "sno" : "1",
        "items" : "Soap",
        "qty" : "2" 
    },
    {
        "sno" : "2",
        "items" : "Dish - Soap",
        "qty" : "1" 
    },
    {
        "sno" : "3",
        "items" : "Paste",
        "qty" : "1" 
    },
]

// Add Data to Sheet
xlWorksheet.addData(data)

// Define a border configuration object
let border = {
    top : 'thin',
    bottom : 'thin',
    left : 'thin',
    right : 'thin',
}

// Set Border
xlWorksheet.border(border)

// Save Workbook
if(xlWorkbook.saveWorkbook('./DemoFile.xlsx')){
    console.log('Workbook has been saved.');
}
else{
    console.log('Workbook cannot be saved..');
}

```
