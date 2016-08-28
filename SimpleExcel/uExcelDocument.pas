// Unit for Easy Using of Excel
unit uExcelDocument;

interface

uses
  System.Generics.Collections, System.IOUtils, Vcl.Dialogs, System.Win.ComObj, System.Classes;

type
  TExcelSheet = class
  private
    FSheet: Variant;
    FName: string;
    FColCount, FRowCount: Integer;
    FTable: TArray<TArray<string>>;
    function GetCell(AX, AY: Integer): string;
    procedure SetCell(AX, AY: Integer; const Value: string);
    function GetName: string;
    procedure SetName(const Value: string);

    constructor Create(AWorkSheet: Variant);
    function GetColCount: Integer;
    function GetRowCount: Integer;
    procedure SetColCount(const Value: Integer);
    procedure SetRowCount(const Value: Integer);
    procedure ApplyToExcel;
  public
    property Cells[AX: Integer; AY: Integer]: string read GetCell write SetCell; default;
    property Name: string read GetName write SetName;
    property RowCount: Integer read GetRowCount write SetRowCount;
    property ColCount: Integer read GetColCount write SetColCount;
    procedure Clear;
    function GetLastFilledLine: Integer;
    destructor Destroy; override;
  end;

  TExcelDocument = class
  private
    FExcel, FWorkBook: Variant;
    FFileName: string;
    FPages: TList<TExcelSheet>;
    function GetPage(Index: Integer): TExcelSheet;
    procedure SetPage(Index: Integer; const Value: TExcelSheet);
    function GetPageCount: Integer;
  public
    constructor CreateFromFile(const AFileName: string);
    constructor CreateNew;
    destructor Destroy; override;
    procedure Save(const AFileName: string);
    procedure Open(const AFileName: string);
    property PageCount: Integer read GetPageCount;
    function AddPage: TExcelSheet;
    property Pages[Index: Integer]: TExcelSheet read GetPage write SetPage;
  end;

implementation

uses
  uExcelUtils, System.SysUtils, uAwaitableClipBoard;

{ TTable }

procedure TExcelSheet.ApplyToExcel;
var
  vX, vY: Integer;
begin
  FSheet.Name := FName;
  for vY := 0 to RowCount - 1 do
    for vX := 0 to ColCount - 1 do
      FSheet.Cells[vY + 1, vX + 1] := FTable[vY, vX];
end;

procedure TExcelSheet.Clear;
var 
  vS: string;
begin
  vS := DecTo26(FColCount) + IntToStr(FRowCount + 1);
  FSheet.Range['a1', vS].Clear;
  FColCount := 1;
  FRowCount := 1;
end;

constructor TExcelSheet.Create(AWorkSheet: Variant);
var
  vS, vOldText: string;
  vStrArr: TArray<string>;
  vStrList: TStringList;
  i: Integer;
  vClipBrd: TAwaitableClipBoard;
begin
  FSheet := AWorkSheet;
  FName := AWorkSheet.Name;
  ColCount := FSheet.UsedRange.Columns.Count;
  RowCount := FSheet.UsedRange.Rows.Count;

  vClipBrd := TAwaitableClipBoard.Create;
  vOldText := vClipBrd.Text;
  
  FSheet.Activate;
  vS := DecTo26(FColCount) + IntToStr(FRowCount + 1);

  FSheet.Range['a1', vS].Select;
  FSheet.Range['a1', vS].Copy;

  vStrList := TStringList.Create;

  vStrList.Text := vClipBrd.Text;
  for i := 0 to vStrList.Count - 1 do
  begin
    vStrArr := vStrList[i].Split([#9]);
    if Length(vStrArr) > 0 then    
      FTable[i] := vStrArr;
  end;
  
  vClipBrd.Text := vOldText;
  vClipBrd.Free;
end;

destructor TExcelSheet.Destroy;
begin
  inherited;
end;

function TExcelSheet.GetCell(AX, AY: Integer): string;
begin
  Result := FTable[AY, AX];
end;

function TExcelSheet.GetColCount: Integer;
begin
  Result := FColCount;
end;

function TExcelSheet.GetLastFilledLine: Integer;
var
  i, j, vW: Integer;
  vS: string;
begin
  vW := FColCount;
  for i := FRowCount - 1 downto 1 do
  begin
    for j := 1 to vW - 1 do
    begin
      vS := FSheet.Cells[i, j];
      if vS <> '' then
        Exit(i);
    end;
  end;
  Result := 0;
end;

function TExcelSheet.GetName: string;
begin
  Result := FName;
end;

function TExcelSheet.GetRowCount: Integer;
begin
  Result := FRowCount;
end;

procedure TExcelSheet.SetCell(AX, AY: Integer; const Value: string);
begin
  FTable[AY, AX]:= Value;
end;

procedure TExcelSheet.SetColCount(const Value: Integer);
var
  i, j, vWidth: Integer;
begin
  vWidth := FColCount;

  if vWidth < Value then
    for i := 0 to FRowCount - 1 do
    begin
      SetLength(FTable[i], Value);
      for j := vWidth to Value - 1 do
        FTable[i, j] := '';
    end;

  if vWidth > Value then
  for i := 0 to FRowCount do
    begin
      for j := FColCount - 1 downto Value do
        FTable[i, j] := '';
      SetLength(FTable[i], Value);
    end;

  FColCount := Value;
end;

procedure TExcelSheet.SetName(const Value: string);
begin
  FName := Value;
end;

procedure TExcelSheet.SetRowCount(const Value: Integer);
var
  i, j, vHeight: Integer;
begin
  vHeight := FRowCount;
  SetLength(FTable, Value);

  if vHeight < Value then
    for i := vHeight to Value - 1 do
    begin
      SetLength(FTable[i], FColCount);
      for j := 0 to FColCount - 1 do
        FTable[i, j] := '';
    end;

  if vHeight > Value then
  for i := vHeight - 1 downto Value do
    begin
      for j := 0 to FColCount - 1 do
        FTable[i, j] := '';
      SetLength(FTable[i], 0);
    end;

  FRowCount := Value;
end;

{ TExcelDocument }

function TExcelDocument.AddPage: TExcelSheet;
var
  vSheet: Variant;
  vN: Integer;
begin
  vN := FWorkBook.Sheets.Count; 
  FExcel.Sheets.Add(After := FWorkBook.Worksheets[vN]);
  vN := FWorkBook.Sheets.Count; 
  vSheet := FExcel.Sheets[vN];

  Result := TExcelSheet.Create(vSheet);
  FPages.Add(Result);
end;

constructor TExcelDocument.CreateFromFile(const AFileName: string);
begin
  FExcel := CreateOleObject('excel.application');
  FPages := TList<TExcelSheet>.Create;
  Open(AFileName);
end;

constructor TExcelDocument.CreateNew;
begin
  FExcel := CreateOleObject('excel.application');
  FExcel.WorkBooks.Add;
  FWorkbook := FExcel.WorkBooks.Item[1];
  FPages := TList<TExcelSheet>.Create;
end;

destructor TExcelDocument.Destroy;
begin
  FExcel.DisplayAlerts := False;
  FExcel.Quit;
  FPages.Free;
  inherited;
end;

function TExcelDocument.GetPage(Index: Integer): TExcelSheet;
begin
  Result := FPages[Index];
end;

function TExcelDocument.GetPageCount: Integer;
begin
  Result := FPages.Count;
end;

procedure TExcelDocument.Open(const AFileName: string);
var
  i: Integer;
begin
  if not TFile.Exists(AFileName) then
  begin
    ShowMessage('File ' + AFileName + ' not found!');
    Exit;
  end;

  FFileName := AFileName;
  FExcel.WorkBooks.Open(AFileName, false);

  FWorkbook := FExcel.WorkBooks.Item[1];

  for i := 1 to FExcel.WorkSheets.Count do
    FPages.Add(TExcelSheet.Create(FWorkbook.Sheets.Item[i]))
end;

procedure TExcelDocument.Save(const AFileName: string);
var
  i: Integer;
begin
  for i := 0 to FPages.Count - 1 do
    FPages[i].ApplyToExcel;

  FWorkBook.SaveAs(AFileName, 51); 
end;

procedure TExcelDocument.SetPage(Index: Integer; const Value: TExcelSheet);
begin
  FPages[Index] := Value;
end;

end.
