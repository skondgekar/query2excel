import Excel from "exceljs";
import { Borders, Font, config } from "./interfaces";
import { Connection, FieldInfo } from "mysql";

export class query2xlsx implements config {
    query: string;
    outFileName: string;
    CellBorders?: Borders;
    HeaderFillColor?: Excel.Fill;
    HeaderFontColor?: Font;
    RowFontColor?: Font;
    RowFillColor?: Excel.Fill;
    connection: Connection;
    freezeFirstColumn?: boolean;
    applyAutoFilter?: boolean;
    private ColumnsWidths: Array<number> = [];

    constructor(input: config) {
        this.query = input.query;
        this.outFileName = input.outFileName;
        this.CellBorders = input.CellBorders;
        this.HeaderFillColor = input.HeaderFillColor;
        this.HeaderFontColor = input.HeaderFontColor;
        this.RowFontColor = input.RowFontColor;
        this.RowFillColor = input.RowFillColor;
        this.connection = input.connection;
        this.freezeFirstColumn = input.freezeFirstColumn;
        this.applyAutoFilter = input.applyAutoFilter;
    }

    getRowFontColor(): Font {
        return {
            name: "Calibri",
            bold: false,
            color: {
                argb: "ff000000"
            }
        }
    }

    getRowFillColor(): Excel.Fill {
        if (this.RowFillColor !== undefined) {
            return this.RowFillColor;
        }
        return {
            fgColor: {
                argb: "ffe3f2fd"
            },
            type: "pattern",
            pattern: "solid"
        };
    }

    getHeaderFillColor(): Excel.Fill {
        if (this.HeaderFillColor !== undefined) {
            return this.HeaderFillColor;
        }
        return {
            fgColor: {
                argb: "ff2196f3"
            },
            type: "pattern",
            pattern: "solid"
        };
    }

    getHeaderFontColor(): Font {
        if (this.HeaderFontColor !== undefined) {
            return this.HeaderFontColor;
        }
        return {
            name: "Calibri",
            bold: true,
            color: {
                argb: "ffffffff"
            }
        }
    }

    getCellBorders(): Borders {
        if (this.CellBorders !== undefined) {
            return this.CellBorders;
        }
        var b: Excel.Border = {
            color: {
                argb: "ffbdbdbd"
            },
            style: "thin"
        }
        return {
            bottom: b,
            top: b,
            left: b,
            right: b
        };
    }
    write(): Promise<any> {
        let p: Promise<any>;
        p = new Promise((resolve, reject) => {
            this.connection.connect((error) => {
                if (error) throw error;
                this.connection.query(this.query, (err, result: Array<any>, fields) => {
                    if (err) {
                        console.log(err);
                    }
                    //Create workbook
                    var wb = new Excel.stream.xlsx.WorkbookWriter({
                        filename: this.outFileName,
                        useStyles: true,
                        useSharedStrings: true
                    });
                    var ws = wb.addWorksheet("Query Result");

                    var output: Array<Array<any>> = [];
                    var numOfColumns = 0;
                    if (fields !== undefined) {
                        var row = ws.getRow(1);

                        fields.forEach((f: FieldInfo, index: number) => {
                            var result = f.name.replace(/([A-Z])/g, " $1");
                            var cell = row.getCell(index + 1);
                            var finalResult = result.charAt(0).toUpperCase() + result.slice(1);
                            cell.value = finalResult.trim();
                            this.ColumnsWidths.push(finalResult.trim().length * 1.5);
                            numOfColumns++;
                        });
                        
                        row.eachCell((cell: Excel.Cell, colNumber: number)=>{
                            cell.fill = this.getHeaderFillColor();
                            cell.font = this.getHeaderFontColor();
                        })

                        ws.columns = fields.map((field, colWidthIndex) => {
                            return {
                                header: field.name,
                                key: field.name,
                                width: this.ColumnsWidths[colWidthIndex]
                            }
                        })
                        row.commit();
                    }
                    var TotalRows = 0;
                    result.forEach((element, index1) => {
                        var row = ws.addRow(element);
                        row.eachCell((cell)=>{
                            cell.fill = this.getRowFillColor();
                            cell.font = this.getRowFontColor();
                            cell.border = this.getCellBorders();
                        })
                        row.commit();
                        TotalRows++;
                    });
                    console.log(numOfColumns);
                    ws.autoFilter = {
                        from: {
                            row: 1,
                            column: 1
                        },
                        to: {
                            row: TotalRows,
                            column: numOfColumns
                        }
                    }

                    // ws.views = [
                    //      { state: 'frozen', ySplit: 1 }
                    // ];
                    ws.commit();
                    wb.commit().then(() => {
                        resolve();
                    })

                });
            })

        });
        return p;
    };
}