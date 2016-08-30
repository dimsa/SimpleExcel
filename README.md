# SimpleExcel
Simple using of Excel (Ole) in Delphi. 

I don't know why, but when I tried to find something simple and reusable for using Excel, I haven't find anything. 

So, here is my wrapper of Excel Ole.

**Example of Creation:**
```
procedure TMainForm.CreateNewBtnClick(Sender: TObject);
var
  vPage: TExcelSheet;
begin
  FExcelDocument := TExcelDocument.CreateNew;
  vPage := FExcelDocument.AddPage;
  vPage.ColCount := 3;
  vPage.RowCount := 3;
  vPage[1, 1] := 'Hello World!';
end;
```
