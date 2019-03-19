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
var workbook = new excel.Workbook();
var worksheet = workbook.addWorksheet('sheet');
var column = [
    { header: 'header_1', key: 'h1', width: 20 },
    { header: 'header_2', key: 'h2', width: 20 },
    { header: 'header_3', key: 'h3', width: 20 }
];
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
var row = worksheet.getRow(1);
var map = new Map();
for (var key in row.model) {
    if (key == 'cells') {
        var model = row.model;
        for (var key_1 in model.cells) {
            if (model.cells.hasOwnProperty(key_1)) {
                var i = parseInt(key_1);
                // console.log(model.cells[i]);   
                map.set(model.cells[i].value, model.cells[i].address);
            }
        }
    }
}
;
console.log(map);
workbook.xlsx.writeFile('./tcr.xlsx');
