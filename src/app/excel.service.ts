import { Injectable } from "@angular/core";
import { Workbook } from "exceljs";
import * as fs from "file-saver";
@Injectable({
  providedIn: "root",
})
export class ExcelService {
  constructor() {}

  generateExcel() {
    //Excel Title, Header, Data
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
    let lengthData = data.length + 3;
    console.log(lengthData);

    //Create workbook and worksheet
    let workbook = new Workbook(); // ws big container ws
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
    // Set name of each cell after merge
    worksheet.getCell("A2").value = "No";
    worksheet.getCell("B2").value = "FULL NAME";
    worksheet.getCell("C2").value = "GENDER";
    worksheet.getCell("D2").value = "SUBJECT";
    worksheet.getCell("G2").value = "AVERAGE SCORE";
    worksheet.getCell("D2").fill = {
      type: "pattern",
      pattern: "solid",
      fgColor: { argb: "FFFFFF00" },
    };
    // Set fill and border for each cell of header
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
    data.forEach((d) => {
      let row = worksheet.addRow(d);
      let qty1 = row.getCell(1);
      let qty2 = row.getCell(2);
      let qty3 = row.getCell(3);
      let qty4 = row.getCell(4);
      let qty5 = row.getCell(5);
      let qty6 = row.getCell(6);
      let qty7 = row.getCell(7);
      qty1.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      qty2.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      qty3.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      qty4.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      qty5.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      qty6.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      qty7.border = {
        left: { style: "thin" },
        right: { style: "thin" },
      };
      let color = "FFFFFFFF";
      if (+qty7.value >= 8) {
        color = "FF99FF99";
      }
      qty1.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
      qty2.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
      qty3.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
      qty4.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
      qty5.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
      qty6.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
      qty7.fill = {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: color },
      };
    });
    worksheet.getColumn(1).alignment = {
      vertical: "middle",
      horizontal: "center",
    };

    worksheet.getColumn(2).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getColumn(4).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getColumn(3).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getColumn(5).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getColumn(6).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getColumn(7).alignment = {
      vertical: "middle",
      horizontal: "center",
    };
    worksheet.getCell(`A:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    worksheet.getCell(`B:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    worksheet.getCell(`C:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    worksheet.getCell(`D:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    worksheet.getCell(`E:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    worksheet.getCell(`F:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };
    worksheet.getCell(`G:${lengthData}`).border = {
      bottom: { style: "thin" },
      left: { style: "thin" },
      right: { style: "thin" },
    };

    worksheet.getColumn(1).width = 5;
    worksheet.getColumn(2).width = 30;
    worksheet.getColumn(3).width = 10;
    worksheet.getColumn(4).width = 10;
    worksheet.getColumn(5).width = 15;
    worksheet.getColumn(6).width = 10;
    worksheet.getColumn(7).width = 30;

    //Generate Excel File with given name
    workbook.xlsx.writeBuffer().then((data) => {
      let blob = new Blob([data], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
      });
      fs.saveAs(blob, "12A6-grade-transcript.xlsx");
    });
  }
}
