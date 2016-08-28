unit uExcelDemo;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants, System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.Grids, Vcl.StdCtrls, uExcelDocument, Vcl.ComCtrls,
  Vcl.ExtCtrls;

type
  TMainForm = class(TForm)
    OpenFileBtn: TButton;
    PageControl: TPageControl;
    SaveFileBtn: TButton;
    SaveDlg: TSaveDialog;
    MainLabel: TLabel;
    CreateNewBtn: TButton;
    ExcelPanel: TPanel;
    AddSheetBtn: TButton;
    AddRowBtn: TButton;
    AddColBtn: TButton;
    OpenDlg: TOpenDialog;
    Button2: TButton;
    procedure OpenFileBtnClick(Sender: TObject);
    procedure SaveFileBtnClick(Sender: TObject);
    procedure AddSheetBtnClick(Sender: TObject);
    procedure AddRowBtnClick(Sender: TObject);
    procedure AddColBtnClick(Sender: TObject);
    procedure FormDestroy(Sender: TObject);
    procedure Button2Click(Sender: TObject);
    procedure CreateNewBtnClick(Sender: TObject);
  private
    FExcelDocument: TExcelDocument;
    procedure OnEnterGrid(Sender: TObject; ACol, ARow: Integer; const Value: string);
    procedure AddSheet(const ASheet: TExcelSheet);
  public

  end;

var
  MainForm: TMainForm;

implementation

{$R *.dfm}

procedure TMainForm.AddColBtnClick(Sender: TObject);
begin
  FExcelDocument.Pages[PageControl.ActivePageIndex].ColCount :=
    FExcelDocument.Pages[PageControl.ActivePageIndex].ColCount + 1;

  TStringGrid(PageControl.ActivePage.Controls[0]).ColCount :=
    FExcelDocument.Pages[PageControl.ActivePageIndex].ColCount;
end;

procedure TMainForm.AddRowBtnClick(Sender: TObject);
begin
  FExcelDocument.Pages[PageControl.ActivePageIndex].RowCount :=
    FExcelDocument.Pages[PageControl.ActivePageIndex].RowCount + 1;

  TStringGrid(PageControl.ActivePage.Controls[0]).RowCount :=
    FExcelDocument.Pages[PageControl.ActivePageIndex].RowCount;
end;

procedure TMainForm.AddSheet(const ASheet: TExcelSheet);
var
  vGrid: TStringGrid;
  vTabSheet: TTabSheet;
  vX, vY: Integer;
begin
  vTabSheet := TTabSheet.Create(PageControl);
  vTabSheet.PageControl := PageControl;

  vGrid := TStringGrid.Create(vTabSheet);
  vGrid.Parent := vTabSheet;
  vGrid.Align := TAlign.alClient;
  vGrid.ColCount := ASheet.ColCount;
  vGrid.RowCount := ASheet.RowCount;
  vGrid.OnSetEditText := OnEnterGrid;
  vGrid.EditorMode := True;
  vGrid.DrawingStyle := TGridDrawingStyle.gdsGradient;
  vGrid.Options := [TGridOption.goEditing, goVertLine, goHorzLine];

  for vX := 0 to ASheet.ColCount - 1 do
    for vY := 0 to ASheet.RowCount - 1 do
      vGrid.Cells[vX, vY] := ASheet.Cells[vX, vY];

  vTabSheet.Caption := ASheet.Name;
end;

procedure TMainForm.AddSheetBtnClick(Sender: TObject);
begin
  AddSheet(FExcelDocument.AddPage);
end;

procedure TMainForm.CreateNewBtnClick(Sender: TObject);
var
  vPage: TExcelSheet;
begin
  FExcelDocument := TExcelDocument.CreateNew;
  vPage := FExcelDocument.AddPage;
  vPage.ColCount := 3;
  vPage.RowCount := 3;
  vPage[1, 1] := 'Hello World!';

  AddSheet(vPage);
end;

procedure TMainForm.Button2Click(Sender: TObject);
var
  i: Integer;
begin
  FExcelDocument := TExcelDocument.CreateFromFile(GetCurrentDir  + '\Test.xlsx');

  for i := 0 to FExcelDocument.PageCount - 1 do
    AddSheet(FExcelDocument.Pages[i]);
end;

procedure TMainForm.FormDestroy(Sender: TObject);
begin
  if Assigned(FExcelDocument) then
    FExcelDocument.Free;
end;

procedure TMainForm.OnEnterGrid(Sender: TObject; ACol, ARow: Integer; const Value: string);
begin
  FExcelDocument.Pages[PageControl.ActivePageIndex].Cells[ACol, ARow] := Value;
end;

procedure TMainForm.OpenFileBtnClick(Sender: TObject);
var
  i: Integer;
begin
  if OpenDlg.Execute then
  begin
    FExcelDocument := TExcelDocument.CreateFromFile(OpenDlg.FileName);

    for i := 0 to FExcelDocument.PageCount - 1 do
      AddSheet(FExcelDocument.Pages[i]);
  end;
end;

procedure TMainForm.SaveFileBtnClick(Sender: TObject);
begin
  if SaveDlg.Execute then
  begin
    FExcelDocument.Save(SaveDlg.FileName);
  end;
end;


end.



