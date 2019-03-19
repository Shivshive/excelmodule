import * as excel from 'exceljs';

let data = [
    { "sno": 1, "item": "IT-1", "qty1": 20, "qty2": 10, "qty3": 10 },
    { "sno": 2, "item": "IT-2", "qty1": 20, "qty2": 10, "qty3": 10 },
    { "sno": 3, "item": "IT-3", "qty1": 20, "qty2": 10, "qty3": 10 },
    { "sno": 4, "item": "IT-4", "qty1": 1, "qty2": 20, "qty3": 30 }
]

const columns = [
    { header: 'SNO', key: 'sno', width: 10 },
    { header: 'Item', key: 'item', width: 30 },
    { header: 'Qty_Jan', key: 'qty1', width: 20 },
    { header: 'Qty_Feb', key: 'qty2', width: 20 },
    { header: 'Qty_March', key: 'qty3', width: 20 },
    { header: 'QtyTotal', key: 'qtytotal', width: 20 }
]


function getRowHeader(worksheet: excel.Worksheet): Map<string, string> {
    // Contains cell information for Row Headers...
    let map = new Map();

    let row = worksheet.getRow(1);

    for (const key in row.model) {
        if (key == 'cells') {
            let model = row.model;

            for (const key in model.cells) {

                if (model.cells.hasOwnProperty(key)) {

                    let i: number = parseInt(key);
                    // console.log(model.cells[i]);   
                    map.set(model.cells[i].value, getColInitials(model.cells[i].address.toString()));
                }
            }
        }
    };

    return map;
}

function getColInitials(colInitial : string){

    return colInitial.substr(0,1);

}

function colTotal(worksheet: excel.Worksheet, no_of_rows_in_col: number , col: string) {

    let colMap = getRowHeader(worksheet);
    let range;
    let firstCell = `${colMap.get(col)}${2}`;
    let lastCell = `${colMap.get(col)}${no_of_rows_in_col+1}`;
    range = `${firstCell}:${lastCell}`;
    return {
        formula : `sum(${range})`
    }
}

const workbook = new excel.Workbook();
const worksheet = workbook.addWorksheet('Invoice');
worksheet.columns = columns;
data.forEach((v, i) => {
    worksheet.addRow(v);
});

worksheet.addRow(
    [
        "QTR Total",
        "",
        colTotal(worksheet,data.length,'Qty_Jan'),
        colTotal(worksheet,data.length,'Qty_Feb'),
        colTotal(worksheet,data.length,'Qty_March')
    ]
)


workbook.xlsx.writeFile('./Formula.xlsx');

