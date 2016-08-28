unit uAwaitableClipBoard;

interface

uses
  Vcl.ClipBrd, System.SysUtils;

type
  TAwaitableClipBoard = class
  private
    function GetText: string;
    procedure SetText(const Value: string);
  public
    property Text: string read GetText write SetText;
    destructor Destroy; override;
  const
    CNAttempts = 10;
  end;

implementation

{ TAwaitableClipBoard }

destructor TAwaitableClipBoard.Destroy;
begin

  inherited;
end;

function TAwaitableClipBoard.GetText: string;
var
  vNAttempts: Integer;
  vText: string;
begin
  vNAttempts := CNAttempts;
  vText := '';
  while vNAttempts > 0 do
  begin
    try
      vText := Clipboard.AsText;
    except
      Sleep(100);
    end;
    Dec(vNAttempts);
  end;
  Result := vText;
end;

procedure TAwaitableClipBoard.SetText(const Value: string);
var
  vNAttempts: Integer;
begin
  vNAttempts := CNAttempts;
  while vNAttempts > 0 do
  begin
    try
      Clipboard.AsText := Value;
    except
      Sleep(100);
    end;
    Dec(vNAttempts);
  end;
end;

end.
