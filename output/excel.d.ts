import * as ex from 'exceljs';
export declare class ExcelWorkbook {
    exWorkbook: ex.Workbook;
    constructor();
    getWorkbook(): ex.Workbook;
    saveWorkbook(path: string): boolean;
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
export declare class ExcelWorksheet {
    exWorkbook: ex.Workbook;
    exsheet: ex.Worksheet;
    sheetName: string;
    options?: sheetConfig;
    constructor(exworkbook: ex.Workbook, sheetName: string, options?: sheetConfig);
    private createWorksheet;
    addHeaders(col: column[]): void;
    addData(data: JSON[]): void;
    colorHeader(headerColour: Partial<mycolour>): void;
    border(style: Border): void;
}
export {};
