unit uProducts;

interface

uses Classes, Generics.Defaults, System.SysUtils;

type

  TRecurringPrice = record
    Price: Double;
    Recurring: Int32;
  end;

  TProduct = class
  public

    Ean: Int64;
    Product: String;
    Category: String;
    Count: Int64;
    Amount: Double;
    RecurringPrice: Double;
    Provider: String;

    constructor Create;

    function SearchObject(ean: Int64; products: TList): TProduct;
    function GetRecurringPrice(prices: System.TArray<Double>): Double;
  end;

implementation

constructor TProduct.Create;
begin
  Ean := 0;
  Product := '';
  Category := '';
  Count := 0;
  Amount := 0.0;
  RecurringPrice := 0.0;
  Provider := '';
end;

function TProduct.SearchObject(ean: Int64; products: TList): TProduct;
var
  product, productSearch: TProduct;
begin

  product := TProduct.Create;

  for productSearch in products do
  begin
    if ean = productSearch.Ean then
    begin
      product := productSearch;
      Break;
    end;
  end;

  Result := product;

end;

function TProduct.GetRecurringPrice(prices: System.TArray<Double>): Double;
var
  LoopA, LoopB, IsExist, MaxCount: Int64;
  priceRecordList: TArray<TRecurringPrice>;
  MaxRecurringPrice: Double;

begin

  SetLength(priceRecordList, 0);

  for LoopA := 0 to Length(prices) -1 do
  begin
    IsExist := 0;

    for LoopB := 0 to Length(priceRecordList) -1 do
    begin

      if priceRecordList[LoopB].Price = prices[LoopA] then
      begin
        Inc(IsExist);
        Inc(priceRecordList[LoopB].Recurring);
      end;

    end;

    if IsExist = 0 then
    begin
      SetLength(priceRecordList, Length(priceRecordList) + 1);
      priceRecordList[Length(priceRecordList) - 1].Price := prices[LoopA];
      priceRecordList[Length(priceRecordList) - 1].Recurring := 1;
    end;

  end;

  MaxCount := 0;
  MaxRecurringPrice := 0.0;

  for LoopB := 0 to Length(priceRecordList) - 1 do
  begin

    if priceRecordList[LoopB].Recurring > MaxCount then
    begin
      MaxRecurringPrice := priceRecordList[LoopB].Price;
      MaxCount := priceRecordList[LoopB].Recurring;
    end;
  end;

  Result := MaxRecurringPrice;

end;

end.
