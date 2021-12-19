import { Component } from "@angular/core";
import { ExcelService } from "./excel.service";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  constructor(private excelService: ExcelService) {}

  generateExcel() {
    const data = [
      [1, "A", 23, "NAM", 5, 7, 6, 5],
      [2, "B", 23, "NAM", 5, 7, 6, 4],
      [3, "C", 23, "NAM", 5, 7, 6, 8],
    ];
    const STYLE_SHEET_TITLE = {
      border: false,
      height: 40,
      font: { size: 30, bold: true, color: { argb: "333333" } },
      alignment: { horizontal: "left", vertical: "middle" },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "ffffff" },
      },
    };
    const STYLE_TABLE_TITLE = {
      border: true,
      height: 40,
      font: { size: 20, bold: false, color: { argb: "333333" } },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "f4f4f4" },
      },
    };
    const STYLE_DATA_BLACK = {
      border: true,
      height: 45,
      font: { size: 15, bold: false, color: { argb: "333333" } },
      alignment: { horizontal: "center", vertical: "middle" },
      fill: {
        type: "pattern",
        pattern: "solid",
        fgColor: { argb: "d1d1d1" },
      },
    };
    const wb = this.excelService.generateWorkbook();
    const sheet1 = this.excelService.addWorksheet(wb, "BẢNG ĐIỂM");
    this.excelService.styleWidthColumns(sheet1, [
      { width: 6 },
      { width: 20 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 15 },
      { width: 25 },
    ]);
    this.excelService.addRowTitle(
      sheet1,
      "HỌC EXCELJS",
      1,
      8,
      STYLE_SHEET_TITLE
    );
    this.excelService.addRowTitle(
      sheet1,
      "BẢNG ĐIỂM HỌC SINH",
      1,
      8,
      STYLE_TABLE_TITLE
    );
    this.excelService.addHeader(
      sheet1,
      this.excelService.STYLE_HEADER,
      [
        "STT",
        "HỌ TÊN",
        "TUỔI",
        "GIỚI TÍNH",
        "ĐIỂM MÔN HỌC",
        "",
        "",
        "ĐIỂM TRUNG BÌNH",
      ],
      ["", "", "", "", "TOÁN", "VĂN", "NGOẠI NGỮ"]
    );
    this.excelService.addData(sheet1, data, STYLE_DATA_BLACK);
    this.excelService.richText(sheet1.getCell("A10"), "helsdfsdfsdfsdflo");
    this.excelService.saveAsExcelFile(wb, "BANG DIEM LOP");
  }
}
