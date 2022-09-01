import { Component, EventEmitter, Input, OnInit, Output } from '@angular/core';
import * as GC from '@grapecity/spread-sheets';
import * as Excel from '@grapecity/spread-excelio';
import { SpreadSheetsModule } from '@grapecity/spread-sheets-angular';

import {saveAs} from 'file-saver';
@Component({
  selector: 'app-sidebar-questions',
  templateUrl: './sidebar-questions.component.html',
  styleUrls: ['./sidebar-questions.component.scss']
})
export class SidebarQuestionsComponent implements OnInit {

  @Input() questions: [];
  @Output() addQuestion = new EventEmitter<any>();
  @Output() deleteQuestion = new EventEmitter<any>();
  @Output() setActiveQuestion = new EventEmitter<any>();
  @Output() importQuestions = new EventEmitter<any>();
  private spread;
  private excelIO;

  
  ngOnInit(): void {
    console.log(this.questions);

  }

  constructor() {
    this.spread = new GC.Spread.Sheets.Workbook();
    this.excelIO = new Excel.IO();
  }

  workbookInit(args: any) {
    const self = this;
    self.spread = args.spread;
    const sheet = self.spread.getActiveSheet();
  
  }

  onFileChange(args: any) {
    console.log(args)
    const self = this, file = args.srcElement && args.srcElement.files && args.srcElement.files[0];
    if (self.spread && file) {
      self.excelIO.open(file, (json: any) => {
        self.spread.fromJSON(json, {});
       console.log(json)
      }, (error: any) => {
        alert('load fail');
      });
    }
  }
  onClickMe(args: any) {
    const self = this;
    const filename = 'exportExcel.xlsx';
    const json = JSON.stringify(self.spread.toJSON());
    self.excelIO.save(json, function (blob: any) {
      //saveAs(blob, filename);
      console.log(json)
    }, function (error: any) {
        console.log(error);
    });
  }

  onAddQuestion() {
    this.addQuestion.emit();
  }

  onImportSpreadSheet(event: any) {
    // this.importQuestions.emit();
   this.onClickMe(event)
  }

  onDeleteQuestion(questionTempId: string) {
    console.log("IN side-que");
    console.log(questionTempId)
    this.deleteQuestion.emit(questionTempId);
  }

  onClickQuestionCard(question) {
    this.setActiveQuestion.emit(question);
  }
}
