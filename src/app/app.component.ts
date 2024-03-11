import { Component, OnInit, ViewChild } from '@angular/core';
import { CoreBase, IUserContext } from '@infor-up/m3-odin';
import { MIService, UserService, ApplicationService } from '@infor-up/m3-odin-angular';
import { SohoFileUploadComponent, SohoTrackDirtyDirective } from 'ids-enterprise-ng';
import * as XLSX from 'xlsx';
@Component({
   selector: 'app-root',
   templateUrl: './app.component.html',
   styleUrls: ['./app.component.css']
})
export class AppComponent extends CoreBase implements OnInit {
   @ViewChild(SohoFileUploadComponent, { static: true }) fileupload?: SohoFileUploadComponent;
   public limitLabel = 'Limited to xls, xlsx and csv';
   public fileUploadOptions = {
      attributes: {
         name: 'data-automation-id',
         value: 'fileupload-field-automation-id'
      }
   };
   public fileUploadDisabled = false;
   public fileUploadReadOnly = false;
   public name1 = 'file-name';
   public fileLimits = '.xls,.xlsx';
   excelData = [];
   userContext = {} as IUserContext;
   isBusy = false;
   company: string;
   currentCompany: string;
   division: string;
   currentDivision: string;
   language: string;
   currentLanguage: string;
   //public val = 0;
   constructor(private miService: MIService, private userService: UserService, private applicationService: ApplicationService) {
      super('AppComponent');
   }

   ngOnInit() {
      this.setBusy(true);
      this.userService.getUserContext().subscribe((userContext: IUserContext) => {
         this.setBusy(false);
         this.logInfo('onClickLoad: Received user context');
         this.userContext = userContext;

      }, (error) => {
         this.setBusy(false);
         this.logError('Unable to get userContext ' + error);
      });
   }


   onClickLoad(): void {
      this.logInfo('onClickLoad');
      this.setBusy(true);
      for (let i = 0; i < this.excelData.length; i++) {
         var year = this.excelData[i].WAYEA4;
         var voucherNo = this.excelData[i].WAVONO;
         var voucherSeries = this.excelData[i].WAVSER;
         var journalNo = this.excelData[i].WBJRNO;
         var journalSequence = this.excelData[i].WBJSNO;
         var text = this.excelData[i].WEVTXT;

         this.applicationService.launch("mforms://_automation?data=%3C%3Fxml%20version%3D%221.0%22%20encoding%3D%22utf-8%22%3F%3E%0A%3Csequence%3E%0A%09%3Cstep%20command%3D%22RUN%22%20value%3D%22GLS200%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAYEA4%22%3E" + year + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVONO%22%3E" + voucherNo + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVSER%22%3E" + voucherSeries + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22LSTOPT%22%20value%3D%225%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WBJRNO%22%3E" + journalNo + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WBJSNO%22%3E" + journalSequence + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22LSTOPT%22%20value%3D%2221%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WEVTXT%22%3E" + text + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%3C%2Fsequence%3E");
         console.log("mforms://_automation?data=%3C%3Fxml%20version%3D%221.0%22%20encoding%3D%22utf-8%22%3F%3E%0A%3Csequence%3E%0A%09%3Cstep%20command%3D%22RUN%22%20value%3D%22GLS200%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAYEA4%22%3E" + year + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVONO%22%3E" + voucherNo + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVSER%22%3E" + voucherSeries + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22LSTOPT%22%20value%3D%225%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WBJRNO%22%3E" + journalNo + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WBJSNO%22%3E" + journalSequence + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22LSTOPT%22%20value%3D%2221%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WEVTXT%22%3E" + text + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%3C%2Fsequence%3E");
      }
      // this.val = 100;
      this.setBusy(false);
      $('body').message({
         title: '<span>M3</span>', status: 'info', message: "Upload coompleted!",
         buttons: [{
            text: 'OK', click: function () { console.log('Info'); $(this).data('modal').close(); }, isDefault: true
         }]
      });
   }

   reverseVoucher(): void {
      this.logInfo('onClickLoad');
      this.setBusy(true);
      for (let i = 0; i < this.excelData.length; i++) {
         var year = this.excelData[i].WAYEA4;
         var voucherNo = this.excelData[i].WAVONO;
         var voucherSeries = this.excelData[i].WAVSER;

         this.applicationService.launch("mforms://_automation?data=%3C%3Fxml%20version%3D%221.0%22%20encoding%3D%22utf-8%22%3F%3E%0A%3Csequence%3E%0A%09%3Cstep%20command%3D%22RUN%22%20value%3D%22GLS200%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAYEA4%22%3E" + year + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVONO%22%3E" + voucherNo + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVSER%22%3E" + voucherSeries + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22LSTOPT%22%20value%3D%2218%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%3C%2Fsequence%3E");
         console.log("mforms://_automation?data=%3C%3Fxml%20version%3D%221.0%22%20encoding%3D%22utf-8%22%3F%3E%0A%3Csequence%3E%0A%09%3Cstep%20command%3D%22RUN%22%20value%3D%22GLS200%22%2F%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAYEA4%22%3E" + year + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVONO%22%3E" + voucherNo + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22AUTOSET%22%3E%0A%09%09%3Cfield%20name%3D%22WAVSER%22%3E" + voucherSeries + "%3C%2Ffield%3E%0A%09%3C%2Fstep%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22LSTOPT%22%20value%3D%2218%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22ENTER%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%09%3Cstep%20command%3D%22KEY%22%20value%3D%22F3%22%2F%3E%0A%3C%2Fsequence%3E");
      }
      // this.val = 100;
      this.setBusy(false);
      $('body').message({
         title: '<span>M3</span>', status: 'info', message: "Upload coompleted!",
         buttons: [{
            text: 'OK', click: function () { console.log('Info'); $(this).data('modal').close(); }, isDefault: true
         }]
      });
   }

   updateUserValues(userContext: IUserContext) {
      this.company = userContext.company;
      this.division = userContext.division;
      this.language = userContext.language;

      this.currentCompany = userContext.currentCompany;
      this.currentDivision = userContext.currentDivision;
      this.currentLanguage = userContext.currentLanguage;
   }

   private setBusy(isBusy: boolean) {
      this.isBusy = isBusy;
   }

   setEnable() {
      (this.fileupload as any).disabled = false;
      this.fileUploadDisabled = (this.fileupload as any).disabled;
      this.fileUploadReadOnly = (this.fileupload as any).readonly;
   }
   setDisable() {
      (this.fileupload as any).disabled = true;
      this.fileUploadDisabled = (this.fileupload as any).disabled;
   }

   onChange(event: any) {
      this.excelData = [];
      console.log('onChange', event);
      const elem: HTMLInputElement = event.currentTarget as HTMLInputElement;
      console.log('file name', elem)
      console.log(elem.files[0].name)
      /* wire up file reader */
      const target: DataTransfer = <DataTransfer>(event.target);
      if (target.files.length !== 1) {
         throw new Error('Cannot use multiple files');
      }
      const reader: FileReader = new FileReader();
      reader.readAsBinaryString(target.files[0]);
      reader.onload = (e: any) => {
         /* create workbook */
         const binarystr: string = e.target.result;
         const wb: XLSX.WorkBook = XLSX.read(binarystr, { type: 'binary' });

         /* selected the first sheet */
         const wsname: string = wb.SheetNames[0];
         const ws: XLSX.WorkSheet = wb.Sheets[wsname];

         /* save data */
         const data = XLSX.utils.sheet_to_json(ws); // to get 2d array pass 2nd parameter as object {header: 1}
         console.log(data); // Data will be logged in array format containing objects
         this.excelData = data;
      };
   }
}
