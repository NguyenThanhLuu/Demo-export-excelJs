import { Component } from "@angular/core";
import { ExcelService } from "./excel.service";
@Component({
  selector: "app-root",
  templateUrl: "./app.component.html",
  styleUrls: ["./app.component.css"],
})
export class AppComponent {
  bodyData: (string | number)[][];
  headers: string[];
  constructor(private excelService: ExcelService) {}
  ngOnInit() {}

  generateExcel() {
    this.excelService.generateExcel();
  }
}
