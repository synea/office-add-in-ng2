import { Component } from '@angular/core';
import { ExcelService } from './office/excel.service';
import { IOfficeResult } from './office/ioffice-result';
//import '../../public/css/styles.css';

@Component({
  selector: 'office-addin-root',
  templateUrl: './app.component.html',
  styleUrls: ['./app.component.css']
})
export class AppComponent {
  constructor(private excelService: ExcelService) { }

  readRange() {
    this.excelService
      .readSelectedRangeData()
      .then((result: IOfficeResult) => {
        console.log(result.data.values);  //this.onResult(result);
      }, (result: IOfficeResult) => {
        console.log(result);  //this.onResult(result);
      });
  }

  getAddress() {
    this.excelService
    .getSelectedRangeAddress()
      .then((result: IOfficeResult) => {
        console.log(result.address);
      }, (result: IOfficeResult) => {
        console.log(result);
      });
  }

}
