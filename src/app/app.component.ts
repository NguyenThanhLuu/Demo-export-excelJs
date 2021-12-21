import { Component } from "@angular/core";
import { ExcelService } from "./excel.service";

@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  constructor(private excelService: ExcelService) {}
  // hàm xử lý lấy file excel có sẵn trong local
  FILE_LIST: any;
  // cần cấu hình style của data ngay từ đầu
  NEW_DATA_CONTENT = [
    [1, "NGUYEN A", 20, "NU", 5, 6, 6, 10],
    [2, "NGUYEN C", 18, "NAM", 5, 9, 6, 10],
    [3, "NGUYEN D", 19, "NU", 5, 4, 6, 8],
  ];
  uploadFile(event: Event) {
    const element = event.currentTarget as HTMLInputElement;
    let fileList: File | null = element.files[0];
    console.log("hi file list: ", fileList);
    this.FILE_LIST = fileList;
    // this.importFilefromLocal(fileList, this.NEW_DATA_CONTENT);
  }
  generateExcel() {
    const data = [
      [1, "A", 23, "NAM", 5, 7, 6, 5],
      [2, "B", 23, "NAM", 5, 7, 6, 4],
      [3, "C", 23, "NAM", 5, 7, 6, 8],
    ];
    const data1 = [
      [1, "D", 22, "NAM", 5, 7, 6, 5, 8],
      [2, "E", 21, "NU", 5, 7, 6, 4, 8],
      [3, "F", 20, "NAM", 5, 7, 6, 8, 10],
    ];
    const cellStart: string = "A2";
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
      "BẢNG ĐIỂM HỌC SINH 7A",
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
    // add và table mới bên dưới table cũ
    this.excelService.addEmptyRow(sheet1, 3);

    this.excelService.addRowTitle(
      sheet1,
      "BẢNG ĐIỂM HỌC SINH 7B",
      1,
      9,
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
        "",
        "ĐIỂM TRUNG BÌNH",
      ],
      ["", "", "", "", "TOÁN", "VĂN", "NGOẠI NGỮ", "HÓA"]
    );
    this.excelService.addData(sheet1, data1, STYLE_DATA_BLACK);

    this.excelService.richText(sheet1.getCell("A10"), "helsdfsdfsdfsdflo");
    this.excelService.saveAsExcelFile(wb, "BANG DIEM LOP");
  }

  exportExcelOverrive() {
    this.importFilefromLocal(this.FILE_LIST, this.NEW_DATA_CONTENT);
  }
  importFilefromLocal(fileImport: any, dataContent: any) {
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
    const DATA_NEW_OVERRIVE = dataContent;
    const wb = this.excelService.generateWorkbook();
    const sheet1 = this.excelService.addWorksheet(wb, "BẢNG ĐIỂM MỚI");
    const reader = new FileReader();
    let dataNew = [];
    reader.readAsArrayBuffer(fileImport);
    reader.onload = () => {
      const buffer: any = reader.result;
      wb.xlsx.load(buffer).then((workbook) => {
        workbook.eachSheet((sheet, id) => {
          sheet.eachRow((row, rowIndex) => {
            console.log(row.values, rowIndex);
            // thay data cũ chỗ này
          });
          sheet.eachRow((row, rowIndex) => {
            for (let i = 5; i < 8; i++) {
              if (rowIndex === i) {
                row.values = DATA_NEW_OVERRIVE[i - 5];
                console.log("cell 1:", row.getCell(1).value);
                // thay đổi style data mới theo style cũ
                for (let x = 1; x < 9; x++) {
                  row.getCell(x).style = {
                    font: { size: 15, bold: false, color: { argb: "333333" } },
                    alignment: { horizontal: "center", vertical: "middle" },
                    fill: {
                      type: "pattern",
                      pattern: "solid",
                      fgColor: { argb: "d1d1d1" },
                    },
                    border: {
                      top: { style: "thin", color: { argb: "858585" } },
                      left: { style: "thin", color: { argb: "858585" } },
                      bottom: { style: "thin", color: { argb: "858585" } },
                      right: { style: "thin", color: { argb: "858585" } },
                    },
                  };
                }
              }
            }
            console.log("kết quả:", rowIndex, row.values);
          });

          this.excelService.saveAsExcelFile(wb, "BANG MOI");
        });
      });
    };
  }
}
