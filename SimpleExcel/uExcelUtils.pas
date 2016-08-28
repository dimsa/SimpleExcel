unit uExcelUtils;

interface

function DecTo26(aValue: Integer): string; // ѕереводит в 26 ричную систему добавл€юю нули до нужной длины


implementation

function NumToEngChar(aValue: Integer): string;
begin
  case aValue of
    1: Result := 'A';
    2: Result := 'B';
    3: Result := 'C';
    4: Result := 'D';
    5: Result := 'E';
    6: Result := 'F';
    7: Result := 'G';
    8: Result := 'H';
    9: Result := 'I';
    10: Result := 'J';
    11: Result := 'K';
    12: Result := 'L';
    13: Result := 'M';
    14: Result := 'N';
    15: Result := 'O';
    16: Result := 'P';
    17: Result := 'Q';
    18: Result := 'R';
    19: Result := 'S';
    20: Result := 'T';
    21: Result := 'U';
    22: Result := 'V';
    23: Result := 'W';
    24: Result := 'X';
    25: Result := 'Y';
    26: Result := 'Z';
    else Result := '';
  end;

end;

function InvertString(aValue: string): string;
var
  i, l: integer;
  vRes: string;
begin
  vRes := '';
  l := Length(aValue);

  for i := l downto 1 do
    vRes := vRes + aValue[i];

  Result := vRes;
end;

function DecTo26(aValue: integer):string; // ѕереводит в 26 ричную систему добавл€юю нули до нужной длины
var
  vA: integer;
  s: string;
begin
  s:='';
  vA := aValue;

  while vA <> 0 do
  begin
    s := s + numToEngChar(vA mod 26);
    vA := trunc(vA / 26);
  end;

  result := InvertString(s);
end;

end.
