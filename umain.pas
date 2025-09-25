unit umain;

interface

uses

  //additional
  ComObj, Winapi.ShellAPI,

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

   ExcelFile : Variant;
   WorkBook : Variant;
   WorkSheet : Variant;
begin
     //be sure ComObj and Variants units are included in the "uses" clause

     ExcelFileName    := ExtractFilePath(Application.ExeName)+'master_xls\Source1.xls';
     ExcelFileNameNew := ExtractFilePath(Application.ExeName)+'output_xls\Target1_'+FormatDateTime('yyyymmdd_hhnnss',Now())+'.xls';

     edSource.text:=ExcelFileName;
     edTarget.Text:=ExcelFileNameNew;

     ExcelApplication := Null;
     ExcelWorkbook := Null;
     ExcelWorksheet := Null;

     try
        ExcelApplication := CreateOleObject('Excel.Application');
     except
           ExcelApplication := Null;
     end;

     //if FileExists(ExcelFileName) then ShowMessage('ada') else ShowMessage('tidak ada');
     

     If VarIsNull(ExcelApplication) = False then
        begin
             try
                ExcelApplication.Visible := False; //set to False if you do not want to see the activity in the background
                ExcelApplication.DisplayAlerts := False; //ensures message dialogs do not interrupt the flow of your automation process. May be helpful to set to True during testing and debugging.

                try
                   ExcelWorkbook := ExcelApplication.Workbooks.Open(ExcelFileName);
                except
                      ExcelWorkbook := Null;
                end;

                If VarIsNull(ExcelWorkbook) = False then
                   begin
                        try
                           ExcelWorksheet := ExcelWorkbook.WorkSheets[1]; //[1] specifies the first worksheet
                        except
                              ExcelWorksheet := Null;
                        end;

                        If VarIsNull(ExcelWorksheet) = False then
                           begin
                                ExcelWorksheet.Select;

                                ExcelWorksheet.Cells[2,3] := edText.Text;                //[row,column], in this case, text is added to Cell A1

//                                ExcelWorkbook.SaveAs(ExcelFileNameNew);
                                //or
                                ExcelApplication.WorkBooks[1].SaveAs(ExcelFileNameNew);
                                //Note: If a file with the new name already exists, it overwrites it. Write additional code to address as desired.
                                //reference
                                //https://docs.microsoft.com/en-us/office/vba/api/excel.workbook.saveas

                                MessageDlg('Proses berhasil !'+#13+#13+'File : '+ExcelFileNameNew, TMsgDlgType.mtInformation, [TMsgDlgBtn.mbOK], 0);

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

  ShellExecute(0, 'open', PChar(ExcelFileNameNew), nil, nil, SW_SHOWNORMAL);

end;

procedure TForm1.FormCreate(Sender: TObject);
begin

  Position  := poScreenCenter;

end;

end.
