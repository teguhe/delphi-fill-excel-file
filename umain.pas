unit umain;

interface

uses

  //additional
  ComObj,

  //autogenerate
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls;

type
  TForm1 = class(TForm)
    lb1: TLabel;
    edText: TEdit;
    btnSendToNewFile: TButton;
    edSource: TEdit;
    edTarget: TEdit;
    procedure FormCreate(Sender: TObject);
    procedure btnSendToNewFileClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  Form1: TForm1;

implementation

{$R *.dfm}

procedure TForm1.btnSendToNewFileClick(Sender: TObject);
var
   ExcelFileName, ExcelFileNameNew: String;
   ExcelApplication, ExcelWorkbook, ExcelWorksheet: Variant;
begin
     //be sure ComObj and Variants units are included in the "uses" clause

     ExcelFileName    := ExtractFilePath(Application.ExeName)+'master_xls\Source1.xls';        //replace file name with the name of your file
     ExcelFileNameNew := ExtractFilePath(Application.ExeName)+'output_xls\Target1_'+FormatDateTime('yyyymmdd_hhnnss',Now())+'.xls'; //replace file name with the name of your file

     edSource.text:=ExcelFileName;
     edTarget.Text:=ExcelFileNameNew;

     ExcelApplication := Null;
     ExcelWorkbook := Null;
     ExcelWorksheet := Null;

     try
        //create Excel OLE
        ExcelApplication := CreateOleObject('Excel.Application');
     except
           ExcelApplication := Null;
           //add error/exception handling code as desired
     end;

     if FileExists(ExcelFileName) then ShowMessage('ada') else ShowMessage('tidak ada');
     

     If VarIsNull(ExcelApplication) = False then
        begin
             try
                ExcelApplication.Visible := True; //set to False if you do not want to see the activity in the background
                ExcelApplication.DisplayAlerts := True; //ensures message dialogs do not interrupt the flow of your automation process. May be helpful to set to True during testing and debugging.
                                             //ShowMessage('1');
                //Open Excel Workbook
                try
                   ExcelWorkbook := ExcelApplication.Workbooks.Open(ExcelFileName);
                   //reference
                   //https://docs.microsoft.com/en-us/office/vba/api/excel.workbooks.open
                except        ShowMessage('2');
                      ExcelWorkbook := Null;
                      //add error/exception handling code as desired
                end;

                If VarIsNull(ExcelWorkbook) = False then
                   begin
                        //connect to Excel Worksheet using either the ExcelApplication or ExcelWorkbook handle
                        try
                           ExcelWorksheet := ExcelWorkbook.WorkSheets[1]; //[1] specifies the first worksheet
                        except
                              ExcelWorksheet := Null;
                              //add error/exception handling code as desired
                        end;

                        If VarIsNull(ExcelWorksheet) = False then
                           begin
                                ExcelWorksheet.Select;

                                ExcelWorksheet.Cells[2,3] := edText.Text;                //[row,column], in this case, text is added to Cell A1
//                                ExcelWorksheet.Cells[2,1] := 125;                                  //adds a number to Cell A2. Cell format determines how Excel will interpret this.
//                                ExcelWorksheet.Cells[3,1] := '=5+5';                               //adds a simple formula to Cell A3
//                                ExcelWorksheet.Cells[4,1] := '=A3*10';                             //adds a simple formula to Cell A4
//                                ExcelWorksheet.Cells[5,1] := '=if(A3=10,TRUE,FALSE)';              //adds a simple formula to Cell A5
//                                ExcelWorksheet.Cells[6,1] := '=if(A4=100,"It matches!","Wrong!")'; //adds a simple formula to Cell A6
//                                ExcelWorksheet.Cells[7,1] := '=today()';                           //adds today's date to Cell A7
//                                ExcelWorksheet.Cells[8,1].ClearContents;                           //clears the contents of Cell A8. It does not clear formatting.
//
//                                ExcelWorksheet.Cells[1,1].Select;                                  //selects Cell A1
                                                     ShowMessage('ok');
//                                ExcelWorkbook.SaveAs(ExcelFileNameNew);
                                //or
                                ExcelApplication.WorkBooks[1].SaveAs(ExcelFileNameNew);
                                //Note: If a file with the new name already exists, it overwrites it. Write additional code to address as desired.
                                //reference
                                //https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.saveas
                           end;
                   end;
             finally
                    ExcelApplication.Workbooks.Close;
                    ExcelApplication.DisplayAlerts := True;
                    ExcelApplication.Quit;

                    ExcelWorksheet := Unassigned;
                    ExcelWorkbook := Unassigned;
                    ExcelApplication := Unassigned;
             end;
        end;
end;

procedure TForm1.FormCreate(Sender: TObject);
begin

  Position  := poScreenCenter;

end;

end.
