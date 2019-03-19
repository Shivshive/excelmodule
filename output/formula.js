"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
var excel = __importStar(require("exceljs"));
var data = [
    { "sno": 1, "item": "IT-1", "qty1": 20, "qty2": 10, "qty3": 10 },
    { "sno": 2, "item": "IT-2", "qty1": 20, "qty2": 10, "qty3": 10 },
    { "sno": 3, "item": "IT-3", "qty1": 20, "qty2": 10, "qty3": 10 },
    { "sno": 4, "item": "IT-4", "qty1": 1, "qty2": 20, "qty3": 30 }
];
var columns = [
    { header: 'SNO', key: 'sno', width: 10 },
    { header: 'Item', key: 'item', width: 30 },
    { header: 'Qty_Jan', key: 'qty1', width: 20 },
    { header: 'Qty_Feb', key: 'qty2', width: 20 },
    { header: 'Qty_March', key: 'qty3', width: 20 },
    { header: 'QtyTotal', key: 'qtytotal', width: 20 }
];
function getRowHeader(worksheet) {
    // Contains cell information for Row Headers...
    var map = new Map();
    var row = worksheet.getRow(1);
    for (var key in row.model) {
        if (key == 'cells') {
            var model = row.model;
            for (var key_1 in model.cells) {
                if (model.cells.hasOwnProperty(key_1)) {
                    var i = parseInt(key_1);
                    // console.log(model.cells[i]);   
                    map.set(model.cells[i].value, getColInitials(model.cells[i].address.toString()));
                }
            }
        }
    }
    ;
    return map;
}
function getColInitials(colInitial) {
    return colInitial.substr(0, 1);
}
function colTotal(worksheet, no_of_rows_in_col, col) {
    var colMap = getRowHeader(worksheet);
    var range;
    var firstCell = "" + colMap.get(col) + 2;
    var lastCell = "" + colMap.get(col) + (no_of_rows_in_col + 1);
    range = firstCell + ":" + lastCell;
    return {
        formula: "sum(" + range + ")"
    };
}
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('Invoice');
worksheet.columns = columns;
data.forEach(function (v, i) {
    worksheet.addRow(v);
});
worksheet.addRow([
    "QTR Total",
    "",
    colTotal(worksheet, data.length, 'Qty_Jan'),
    colTotal(worksheet, data.length, 'Qty_Feb'),
    colTotal(worksheet, data.length, 'Qty_March')
]);
workbook.xlsx.writeFile('./Formula.xlsx');
