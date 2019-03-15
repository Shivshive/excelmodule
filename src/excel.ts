import * as ex from 'exceljs'

export class ExcelWorkbook {

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

interface sheetConfig {
    properties: Partial<worksheetProperties>;
}

interface worksheetProperties {
    tabColor: mycolour;
    showGridLines: boolean;
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

export class ExcelWorksheet {

    exWorkbook : ex.Workbook;
    exsheet!: ex.Worksheet;
    sheetName: string;
    options?: sheetConfig;

    constructor(exworkbook : ex.Workbook, sheetName: string, options?: sheetConfig) {

        this.exWorkbook = exworkbook; 
        this.sheetName = sheetName;
        if (options) {
            this.options = options;
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
