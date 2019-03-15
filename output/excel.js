"use strict";
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
var ex = __importStar(require("exceljs"));
var ExcelWorkbook = /** @class */ (function () {
    function ExcelWorkbook() {
        this.exWorkbook = new ex.Workbook();
    }
    ExcelWorkbook.prototype.getWorkbook = function () {
        return this.exWorkbook;
    };
    ExcelWorkbook.prototype.saveWorkbook = function (path) {
        try {
            this.exWorkbook.xlsx.writeFile(path);
            return true;
        }
        catch (e) {
            return false;
        }
    };
    return ExcelWorkbook;
}());
exports.ExcelWorkbook = ExcelWorkbook;
var ExcelWorksheet = /** @class */ (function () {
    function ExcelWorksheet(exworkbook, sheetName, options) {
        this.exWorkbook = exworkbook;
        this.sheetName = sheetName;
        if (options) {
            this.options = options;
        }
        this.exsheet = this.createWorksheet();
    }
    ExcelWorksheet.prototype.createWorksheet = function () {
        if (this.options) {
            return this.exWorkbook.addWorksheet(this.sheetName, this.options);
        }
        else {
            return this.exWorkbook.addWorksheet(this.sheetName);
        }
    };
    ExcelWorksheet.prototype.addHeaders = function (col) {
        this.exsheet.columns = col;
    };
    ExcelWorksheet.prototype.addData = function (data) {
        this.exsheet.addRows(data);
    };
    ExcelWorksheet.prototype.colorHeader = function (headerColour) {
        var row = this.exsheet.getRow(1);
        row.eachCell(function (c, i) {
            c.fill = {
                type: 'pattern',
                pattern: 'solid',
                fgColor: headerColour,
                bgColor: headerColour
            };
        });
    };
    ExcelWorksheet.prototype.border = function (style) {
        this.exsheet.eachRow(function (row) {
            row.eachCell({ includeEmpty: false }, function (cell) {
                cell.border = style;
            });
        });
    };
    return ExcelWorksheet;
}());
exports.ExcelWorksheet = ExcelWorksheet;
