unit uCsv;

interface

uses System.Classes, System.SysUtils;

type
  TCSVRecords = class
  public

    Ean: Int64;
    Product: String;
    Category: String;
    Provider: String;
    Invoice: Int64;
    Count: Int64;
    Amount: Double;
    Price: Double;

    Constructor Create;

    procedure SetPrice();
    function CheckRowIntegrity(rows: TStrings): Boolean;
    function CheckColIntegrity(columns: TStrings): Boolean;
  end;

implementation

Constructor TCSVRecords.Create;
begin
  Ean := 0;
  Product := '';
  Category := '';
  Provider := '';
  Invoice := 0;
  Count := 0;
  Amount := 0.0;
end;

procedure TCSVRecords.SetPrice;
begin
  if Count > 0 then Price := Amount / Count
  else Price := 0;
end;

function TCSVRecords.CheckRowIntegrity(rows: TStrings): Boolean;
var
  ColsOne, ColsTwo: TStrings;
begin

  Result := True;

  if Not rows.Count > 1 then
  begin
    Result := False;
  end
  else
  begin
    ColsOne := TStringList.Create;
    ColsOne.Delimiter := ';';
    ColsOne.StrictDelimiter := True;
    ColsOne.DelimitedText := Rows[0];

    if ColsOne.Count <> 11 then
    begin
      Result := False;
    end
    else
    begin

      ColsTwo := TStringList.Create;
      ColsTwo.Delimiter := ';';
      ColsTwo.StrictDelimiter := True;
      ColsTwo.DelimitedText := Rows[1];

      if CheckColIntegrity(ColsTwo) = False then Result := False;
    end;
  end;
end;

function TCSVRecords.CheckColIntegrity(columns: TStrings): Boolean;
var
  IntTest: Int64;
  DoubleTest: Double;
begin

  Result := True;

  if columns.Count <> 11 then Result := False;
  if TryStrToInt64(columns[0], IntTest) = False then Result := False;
  if TryStrToInt64(columns[4], IntTest) = False then Result := False;
  if TryStrToInt64(columns[5], IntTest) = False then Result := False;
  if TryStrToFloat(columns[6].Replace('.', ''), DoubleTest) = False then Result := False;

end;

end.
