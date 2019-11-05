unit UMain;

interface

uses
  Windows, Messages, SysUtils, Variants, Classes, Graphics, Controls, Forms,
  Dialogs, StdCtrls, Grids, DBGrids, ExtCtrls, DB, ADODB, frxClass, frxCross;

type
  TFrmMain = class(TForm)
    pnl1: TPanel;
    Dbgrd_Excel: TDBGrid;
    pnl2: TPanel;
    pnl3: TPanel;
    Edt_ExcelFilePath: TEdit;
    Btn_OpenFile: TButton;
    CBSheet: TComboBox;
    Btn_LoadExcel: TButton;
    Label1: TLabel;
    Label2: TLabel;
    Btn_Exit: TButton;
    Label3: TLabel;
    AdoCon_Excel: TADOConnection;
    AdoQuery_Excel: TADOQuery;
    DlgOpen_Excel: TOpenDialog;
    ds_Excel: TDataSource;
    Btn_Print: TButton;
    frxReport1: TfrxReport;
    procedure Btn_ExitClick(Sender: TObject);
    procedure Btn_OpenFileClick(Sender: TObject);
    procedure Btn_LoadExcelClick(Sender: TObject);
    procedure Btn_PrintClick(Sender: TObject);
    procedure Before_Print(c: TfrxReportComponent);
  private
    { Private declarations }
    procedure Connect_To_Excel;
    procedure Get_Excel_Data;
    function get_Template_Dir: string;
    function get_Template_FullPath: string;
  public
    { Public declarations }
  end;

var
  FrmMain: TFrmMain;

implementation
  uses
     Vcl.OleAuto, AdjustGrid, System.UITypes, System.IOUtils;
{$R *.dfm}

procedure TFrmMain.Connect_To_Excel;
var
  strConn:  widestring;
begin
  strConn:='Provider=Microsoft.Jet.OLEDB.4.0;' +
           'Data Source=' + Edt_ExcelFilePath.Text + ';' +
           'Extended Properties=Excel 8.0;';

  AdoCon_Excel.Connected:=False;
  AdoCon_Excel.ConnectionString := strConn;
  try
    AdoCon_Excel.Open;
    AdoCon_Excel.GetTableNames(CBSheet.Items,True);
  except
    MessageDlg('Failed To Connect to Excel File !!',mtWarning,[mbok],0);
  end;
end;

function TFrmMain.get_Template_Dir: string;
begin
  result :=  TPath.GetDirectoryName(TPath.GetFullPath(paramstr(0)))+'\report';
end;

function TFrmMain.get_Template_FullPath: string;
begin
  result :=  TPath.GetDirectoryName(TPath.GetFullPath(paramstr(0)))+'\report\Excel.fr3';
end;

procedure TFrmMain.Get_Excel_Data;
begin
  if not AdoCon_Excel.Connected then Connect_To_Excel;
  AdoQuery_Excel.Close;
  AdoQuery_Excel.SQL.Text:='SELECT * FROM ['+CBSheet.Text+']';
  try
    AdoQuery_Excel.Open;
  except
    MessageDlg('Failed To Connect to Excel File !!',mtWarning,[mbok],0);
  end;
  if not AdoQuery_Excel.IsEmpty then
  begin
    Btn_Print.Enabled := true;
  end;

end;

procedure TFrmMain.Before_Print(c: TfrxReportComponent);
var
  Cross: TfrxCrossView;
  i, j: Integer;
begin
  if c is TfrxCrossView then
  begin
    Cross := TfrxCrossView(c);

    AdoQuery_Excel.First;
    i := 0;
    while not AdoQuery_Excel.Eof do
    begin
      for j := 0 to AdoQuery_Excel.Fields.Count - 1 do
        Cross.AddValue([i], [AdoQuery_Excel.Fields[j].DisplayLabel], [AdoQuery_Excel.Fields[j].AsString]);

      AdoQuery_Excel.Next;
      Inc(i);
    end;
  end;
end;

procedure TFrmMain.Btn_ExitClick(Sender: TObject);
begin
  Close;
end;

procedure TFrmMain.Btn_OpenFileClick(Sender: TObject);
begin
  DlgOpen_Excel.InitialDir := ExtractFileDir(ParamStr(0));

  if DlgOpen_Excel.Execute then
  begin
    Edt_ExcelFilePath.Text := DlgOpen_Excel.FileName;
    Connect_To_Excel;
    CBSheet.ItemIndex := 0;
    Btn_LoadExcel.Enabled := True;
  end else ShowMessage('please choose a Simple Excel File to load it here');
end;

procedure TFrmMain.Btn_PrintClick(sender: TObject);
begin
  frxReport1.LoadFromFile(get_Template_FullPath);
  frxReport1.ShowReport;
  Btn_Print.Enabled := false;
end;

procedure TFrmMain.Btn_LoadExcelClick(Sender: TObject);
begin
  Get_Excel_Data; AdjustColumnWidths(Dbgrd_Excel);
end;

end.
