import { Injectable } from '@angular/core';
import { IOfficeResult } from './ioffice-result';

@Injectable()
export class ExcelService {

    constructor() { }

    //Returns address sting
    getSelectedRangeAddress(): Promise<IOfficeResult> {
        return new Promise((resolve, reject) => {
            Excel.run((ctx: Excel.RequestContext) => {
                let rangeData = ctx.workbook.getSelectedRange().load();
                return ctx.sync().then(() => {                    
                    return rangeData;
                });
            }).then(data => {
                resolve({ 'address': data.address });
            }, error => {
                reject({ 'error': error.message });
            });
        });
    }

    readSelectedRangeData(): Promise<IOfficeResult> {
        return new Promise((resolve, reject) => {
            Excel.run((ctx: Excel.RequestContext) => {
                let rangeData = ctx.workbook.getSelectedRange().load("values");
                return ctx.sync().then(() => {
                    // TO DO: return error if rangeData.values.length == 0  Meaning function got called with no selection made
                    return rangeData;
                });
            }).then(data => {
                resolve({ 'data': data });
            }, error => {
                reject({ 'error': error.message });
            });
        });
    }
}