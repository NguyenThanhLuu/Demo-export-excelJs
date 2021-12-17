import { Injectable } from "@angular/core";
import { Workbook } from "exceljs";
import * as fs from "file-saver";
@Injectable({
  providedIn: "root",
})
export class ExcelService {
  constructor() {}
  title = "12A6 GRADE TRANSCRIPT";
  generateExcel() {
    // Excel Title, Header, Data
    const title = "12A6 GRADE TRANSCRIPT";
    const header = ["", "", "", "MATH", "LITERATURE", "ENGLISH", ""];
    const data = [
      [1, "PHAM NHU ANH", "Female", 9, 10, 10, 9.7],
      [2, "PHAM THANH BINH", "	Male	", 8, 8, 9, 8.3],
      [3, "TRAN QUOC BINH", "	Male", 8, 9, 8, 8.3],
      [4, "LE THI HOAI AN	", "Female", 8, 8, 8, 8.0],
      [5, "DANG QUOC AN	", "Male", 8, 9, 7, 8.0],
      [6, "HOANG DUNG	", "Male", 8, 8, 7, 7.7],
      [7, "LE XUAN TRUONG", "	Male", 8, 8, 8, 8.0],
      [8, "BUI TIEN DUNG", "	Male", 8, 8, 6, 7.3],
      [9, "DANG THI THUY HOA", "Female", 7, 9, 6, 7.3],
      [10, "NGUYEN TIEN DUNG", "	Male", 10, 9, 9, 9.3],
      [11, "NGUYEN QUANG HAI	", "Male	", 8, 8, 7, 7.7],
      [12, "PHAN VAN DUC", "Male", 8, 7, 7, 7.3],
    ];
    let valueMergeHeader = [
      "No",
      "FULL NAME",
      "GENDER",
      "SUBJECT",
      "AVERAGE SCORE",
    ];
    let positionCell = ["A2", "B2", "C2", "D2", "G2"];

    let lengthData = data.length + 3;
    console.log(lengthData);

    // Function set title and value
    function setTitle(position: string[], valueCell: string[]) {
      for (let i = 0; i < position.length; i++) {
        worksheet.getCell(position[i]).value = valueCell[i];
      }
    }
    // Function fill
    function fillCell(nameCell: string, colorFill: string) {
      worksheet.getCell(nameCell).fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: colorFill },
      };
    }
    // Function border
    function borderCell(nameCell: string) {
      worksheet.getCell(nameCell).border = {
        bottom: { style: "thin" },
        top: { style: "thin" },
        left: { style: "thin" },
        right: { style: "thin" },
      };
    }
    // Function align text in a cell (center)
    function alignColumn(numberColumn: number) {
      worksheet.getColumn(numberColumn).alignment = {
        vertical: "middle",
        horizontal: "center",
      };
    }
    // Function width column
    function widthColumn(numberColumn: number, valueSet: number) {
      worksheet.getColumn(numberColumn).width = valueSet;
    }
    // Create workbook and worksheet
    let workbook = new Workbook();
    let worksheet = workbook.addWorksheet("12A6 grade transcript");

    // Add title and excuse set font, background cell and border
    let titleRow = worksheet.addRow([title]);
    // Set font for title
    titleRow.font = {
      name: "consolas",
      family: 4,
      size: 16,
      bold: true,
      color: { argb: "FFFFFF" },
    };
    // Set text posision in cell
    titleRow.alignment = { vertical: "middle", horizontal: "center" };
    titleRow.findCell(1).fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "6600ff" }, // Purple color
    };
    // Set border of cell
    titleRow.findCell(1).border = {
      top: { style: "thin" },
      left: { style: "thin" },
      bottom: { style: "thin" },
      right: { style: "thin" },
    };
    // Merge cell for title
    worksheet.mergeCells("A1:G1");
    // Add Header Row
    worksheet.addRow([]);
    let headerRow = worksheet.addRow(header);
    // Merge cell of header
    worksheet.mergeCells("A2:A3");
    worksheet.mergeCells("B2:B3");
    worksheet.mergeCells("C2:C3");
    worksheet.mergeCells("D2:F2");
    worksheet.mergeCells("G2:G3");

    setTitle(positionCell, valueMergeHeader);

    // Set fill and border for each cell of header
    fillCell("D2", "FFFFFF00");
    headerRow.eachCell((cell, number) => {
      cell.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "FFFFFF00" }, // Yellow color
      };
      cell.border = {
        top: { style: "thin" },
        left: { style: "thin" },
        bottom: { style: "thin" },
        right: { style: "thin" },
      };
    });
    // Add each row corresponding to each value of data array
    data.forEach((value) => {
      let row = worksheet.addRow(value);
      for (let i = 1; i <= 7; i++) {
        row.getCell(i).border = {
          left: { style: "thin" },
          right: { style: "thin" },

          bottom: { style: "thin" },
          top: { style: "thin" },
        };
        if (i == 7 && +row.getCell(i).value >= 8) {
          row.getCell(i).fill = {
            type: "pattern",
            pattern: "solid",
            fgColor: { argb: "b3b3b3" },
          };
          console.log(1);
        }
      }
    });
    // Alignment text in cell
    for (let i = 1; i <= 7; i++) {
      alignColumn(i);
    }
    // Modify with each cell
    widthColumn(1, 5);
    widthColumn(2, 30);
    widthColumn(3, 10);
    widthColumn(4, 10);
    widthColumn(5, 15);
    widthColumn(6, 10);
    widthColumn(7, 30);

    // Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      fs.saveAs(blob, "12A6-grade-transcript.xlsx");
    });
  }
}
