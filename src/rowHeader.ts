import * as excel from 'exceljs';
import { stringify } from 'querystring';

let workbook = new excel.Workbook();
let worksheet = workbook.addWorksheet('sheet');

let column = [
    { header: 'header_1', key: 'h1', width: 20 },
    { header: 'header_2', key: 'h2', width: 20 },
    { header: 'header_3', key: 'h3', width: 20 }
]

worksheet.columns = column;

worksheet.addRows([
    {
        "h1": "1",
        "h2": "2",
        "h3": "2"
    },
    
    {
        "h1": "1",
        "h2": "1",
        "h3": "3"
    }

]);

let row = worksheet.getRow(1);

// console.log(row.model);


interface cellModel {
    address : Address;
    type : excel.ValueType;
    value : excel.CellValue;
}


interface Address{
    address : string;
}

let map = new Map();

for (const key in row.model) {

    if(key == 'cells'){
        let model = row.model;
        for (const key in model.cells) {
            if (model.cells.hasOwnProperty(key)) {
                let i : number = parseInt(key);
                // console.log(model.cells[i]);   
                map.set(model.cells[i].value, model.cells[i].address);
            }
        }
    }
};

console.log(map);
workbook.xlsx.writeFile('./tcr.xlsx');
