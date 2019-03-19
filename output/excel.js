"use strict";
var __assign = (this && this.__assign) || function () {
    __assign = Object.assign || function(t) {
        for (var s, i = 1, n = arguments.length; i < n; i++) {
            s = arguments[i];
            for (var p in s) if (Object.prototype.hasOwnProperty.call(s, p))
                t[p] = s[p];
        }
        return t;
    };
    return __assign.apply(this, arguments);
};
var __importStar = (this && this.__importStar) || function (mod) {
    if (mod && mod.__esModule) return mod;
    var result = {};
    if (mod != null) for (var k in mod) if (Object.hasOwnProperty.call(mod, k)) result[k] = mod[k];
    result["default"] = mod;
    return result;
};
Object.defineProperty(exports, "__esModule", { value: true });
var ex = __importStar(require("exceljs"));
function rowTotal(value, worksheet) {
    var cols = [];
    for (var _i = 2; _i < arguments.length; _i++) {
        cols[_i - 2] = arguments[_i];
    }
    var rowHeaders = getRowHeader(worksheet);
    // console.log(` Row index : ${value}`)
    var colAddArray = [];
    cols.forEach(function (colHeader) {
        // console.log(colHeader);
        var cellAddress = rowHeaders.get(colHeader);
        // console.log(`Column ${colHeader} has cell address of ${cellAddress}`);
        colAddArray.push("" + cellAddress);
    });
    var range = "";
    if (colAddArray) {
        var falphabet = colAddArray[0].substr(0, 1);
        var fnumber = colAddArray[0].substr(1, 1);
        var firstCellReference = "" + falphabet + (parseInt(fnumber) + value + 1);
        var lalphabet = colAddArray[colAddArray.length - 1].substr(0, 1);
        var lnumber = colAddArray[colAddArray.length - 1].substr(1, 1);
        var lastCellReference = "" + lalphabet + (parseInt(lnumber) + value + 1);
        range = firstCellReference + " : " + lastCellReference;
        // console.log(range);
    }
    return {
        formula: "sum(" + range + ")"
    };
}
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
                    map.set(model.cells[i].value, model.cells[i].address);
                }
            }
        }
    }
    ;
    return map;
}
var exWorkbook = /** @class */ (function () {
    function exWorkbook() {
        this.exWorkbook = new ex.Workbook();
    }
    exWorkbook.prototype.getWorkbook = function () {
        return this.exWorkbook;
    };
    exWorkbook.prototype.saveWorkbook = function (path) {
        try {
            this.exWorkbook.xlsx.writeFile(path);
            return true;
        }
        catch (e) {
            return false;
        }
    };
    return exWorkbook;
}());
exports.exWorkbook = exWorkbook;
var exWorksheet = /** @class */ (function () {
    function exWorksheet(exworkbook, sheetName, options) {
        this.exWorkbook = exworkbook;
        this.sheetName = sheetName;
        if (options) {
            this.options = options;
            console.log(this.options);
        }
        this.exsheet = this.createWorksheet();
    }
    exWorksheet.prototype.createWorksheet = function () {
        if (this.options) {
            return this.exWorkbook.addWorksheet(this.sheetName, this.options);
        }
        else {
            return this.exWorkbook.addWorksheet(this.sheetName);
        }
    };
    exWorksheet.prototype.addHeaders = function (col) {
        this.exsheet.columns = col;
    };
    exWorksheet.prototype.addData = function (data) {
        this.exsheet.addRows(data);
    };
    exWorksheet.prototype.addDataWithRowTotal = function (data, totalHeader) {
        var _this = this;
        var col = [];
        for (var _i = 2; _i < arguments.length; _i++) {
            col[_i - 2] = arguments[_i];
        }
        console.log.apply(console, col);
        data.forEach(function (v, i) {
            var _a;
            _this.exsheet.addRow(__assign({}, v, (_a = {}, _a[totalHeader] = rowTotal.apply(void 0, [i, _this.exsheet].concat(col)), _a)));
        });
    };
    exWorksheet.prototype.colorHeader = function (headerColour) {
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
    exWorksheet.prototype.border = function (style) {
        this.exsheet.eachRow(function (row) {
            row.eachCell({ includeEmpty: false }, function (cell) {
                cell.border = style;
            });
        });
    };
    return exWorksheet;
}());
exports.exWorksheet = exWorksheet;
