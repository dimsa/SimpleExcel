program ExcelDemo;

uses
  Vcl.Forms,
  uExcelDemo in 'uExcelDemo.pas' {MainForm},
  uAwaitableClipBoard in 'SimpleExcel\uAwaitableClipBoard.pas',
  uExcelDocument in 'SimpleExcel\uExcelDocument.pas',
  uExcelUtils in 'SimpleExcel\uExcelUtils.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  Application.CreateForm(TMainForm, MainForm);
  Application.Run;
end.
