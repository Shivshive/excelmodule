import * as ex from 'exceljs'

interface sheetConfig {
    properties: Partial<worksheetProperties>;
    views : Array<Partial<worksheetViews>>;
}

interface worksheetProperties {
    tabColor: mycolour;
    showGridLines: boolean;
}

interface worksheetViews{
    showRuler : boolean;
    showGridLines : boolean;
}


interface mycolour {
    argb: string;
}

interface column {
    header: string;
    key: string;
    width: number;
}

interface BorderStyle {
    style: Partial<ex.BorderStyle>;
}

interface Border {
    top: Partial<BorderStyle>;
    bottom: Partial<BorderStyle>;
    left: Partial<BorderStyle>;
    right: Partial<BorderStyle>;
}

function rowTotal(value: number, worksheet: ex.Worksheet , ...cols : string[]) {

    let rowHeaders = getRowHeader(worksheet);
    // console.log(` Row index : ${value}`)
    let  colAddArray : string[] = [];

    cols.forEach(colHeader => {
        // console.log(colHeader);
        let cellAddress = rowHeaders.get(colHeader);
        // console.log(`Column ${colHeader} has cell address of ${cellAddress}`);
        colAddArray.push(`${cellAddress}`);
    });
    
    let range : string = "";

    if(colAddArray){
        let falphabet = colAddArray[0].substr(0,1);
        let fnumber = colAddArray[0].substr(1,1);
        let firstCellReference = `${falphabet}${parseInt(fnumber) + value + 1}`;

        let lalphabet = colAddArray[colAddArray.length-1].substr(0,1);
        let lnumber = colAddArray[colAddArray.length-1].substr(1,1);
        let lastCellReference = `${lalphabet}${parseInt(lnumber) + value + 1}`;
        range = `${firstCellReference} : ${lastCellReference}`;
        // console.log(range);
    }
 
    return {
        formula: `sum(${range})`
    }
}

function getColInitials(colInitial : string){

    return colInitial.substr(0,1);

}

function colTotal(worksheet: ex.Worksheet, no_of_rows_in_col: number , col: string) {

    let colMap = getRowHeader(worksheet);
    let range;
    let firstCell = `${colMap.get(col)}${2}`;
    let lastCell = `${colMap.get(col)}${no_of_rows_in_col+1}`;
    range = `${firstCell}:${lastCell}`;
    return {
        formula : `sum(${range})`
    }
}

function getRowHeader(worksheet: ex.Worksheet): Map<string, string> {
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


export class exWorkbook {

    exWorkbook: ex.Workbook;
    
    constructor() {
        this.exWorkbook = new ex.Workbook();
    }

    getWorkbook() : ex.Workbook {
        return this.exWorkbook;
    }

    saveWorkbook(path: string): boolean {
        try {
            this.exWorkbook.xlsx.writeFile(path);
            return true;
        } catch (e) {
            return false;
        }
    }

}
export class exWorksheet {

    exWorkbook : ex.Workbook;
    exsheet!: ex.Worksheet;
    sheetName: string;
    options?: sheetConfig;

    constructor(exworkbook : ex.Workbook, sheetName: string, options?: sheetConfig) {

        this.exWorkbook = exworkbook; 
        this.sheetName = sheetName;
        if (options) {
            this.options = options;
            console.log(this.options);
        }
        this.exsheet = this.createWorksheet();
    }

    private createWorksheet() : ex.Worksheet {
        if (this.options) {
            return this.exWorkbook.addWorksheet(this.sheetName, this.options);
        } else {
            return this.exWorkbook.addWorksheet(this.sheetName);
        }

    }

    addHeaders(col: column[]) {
        this.exsheet.columns = col;
    }

    addData(data: JSON[]) {
        this.exsheet.addRows(data);
    }

    addDataWithRowTotal(data : JSON[], totalHeader : string , ...col : string[]){
        console.log(...col);

        data.forEach((v,i)=>{
            this.exsheet.addRow({
                ...v,
                [totalHeader] : rowTotal(i,this.exsheet, ...col)
            })
        })
    }

    colorHeader(headerColour: Partial<mycolour>) {
        let row = this.exsheet.getRow(1);
        row.eachCell((c, i) => {
            c.fill = {
                type : 'pattern',
                pattern : 'solid',
                fgColor : headerColour,
                bgColor : headerColour
            }
        });
    }

    border(style: Border) {
        this.exsheet.eachRow((row) => {
            row.eachCell({ includeEmpty: false }, (cell) => {
                cell.border = style;
            })
        })
    }
}
