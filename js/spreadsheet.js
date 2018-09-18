/**
 * This file gets compiled as bundle-appexcel.js by browserify in order
 * to work in the browser without a node.js server. Also uses file-saverjs to
 * make use of file saving.
 * @author Carlos Eli Lopez Tellez
 */

var Filesaverjs = require('file-saverjs');
var Excel       = require('exceljs/dist/es5/exceljs.browser');

global.Spreadsheet = function Spreadsheet(allData) {
    this.workbook;
    this.worksheet;
    this.fileName;
    this.allData = allData;

    this.generate = () => {
        try {
            this.workbook = new Excel.Workbook();
            this.worksheet = this.workbook.addWorksheet('1');
        } catch(err) {
            console.log('Error creating spreadsheet file: ' + err);
        }
    }


    this.setStaticCells = (consistent) => {
        this.fileName = consistent.title;
        
        this.worksheet.mergeCells('A1:D1');
        
        var cell       = this.worksheet.getCell('D1');
        cell.value     = consistent.title;
        cell.font      = { bold: true };
        cell.alignment = { vertical: 'middle', horizontal: 'center' };
        cell.border    = { top: { style:'thick', color: { argb:'000000' } },
            left:   { style:'thick', color:{argb:'000000'} },
            bottom: { style:'thick', color:{argb:'000000'} },
            right:  { style:'thick', color:{argb:'000000'} }
        };
 
        this.worksheet.mergeCells('B2:D2');
        
        
        cell = this.worksheet.getCell('D2');
        cell.font = { size: 8 };
        cell.alignment = {wrapText: true};
        cell.value = consistent.disclosure;
    }

    this.setDimensions = (dimensions) => {
        for (const index in dimensions.columnWidths) {
            const column = this.worksheet.getColumn(index);
            column.width = dimensions.columnWidths[index];
        }

        for (const index in dimensions.rowHeights) {
            const row = this.worksheet.getRow(index);
            row.height = dimensions.rowHeights[index];
        }
    }

    this.downloadFile = () => {
        const fileName = this.fileName;
        
        this.workbook.xlsx.writeBuffer().then((buffer) => { 
            Filesaverjs(new Blob([buffer],{type:"application/octet-stream"}), fileName+".xlsx"); 
        }); 
    }
}