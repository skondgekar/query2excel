import Excel from "exceljs";
import { Connection } from "mysql";

export interface config {
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
}

export interface Borders {
    bottom: Excel.Border,
    top: Excel.Border,
    left: Excel.Border,
    right: Excel.Border
}

export interface Font {
	name: string;
	size?: number;
	family?: number;
	scheme?: 'minor' | 'major' | 'none';
	charset?: number;
	color: Partial<Color>;
	bold?: boolean;
	italic?: boolean;
	underline?: boolean | 'none' | 'single' | 'double' | 'singleAccounting' | 'doubleAccounting';
	vertAlign?: 'superscript' | 'subscript';
	strike?: boolean;
	outline?: boolean;
}

export interface Color {
	/**
	 * Hex string for alpha-red-green-blue e.g. FF00FF00
	 */
	argb: string;

}


