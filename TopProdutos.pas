unit TopProdutos;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.StdCtrls, Vcl.Imaging.pngimage,
  Vcl.ExtCtrls, Vcl.Buttons, Vcl.Grids, Vcl.ComCtrls, Data.DB, Vcl.DBGrids,
  Generics.Defaults;

type

  TExclusiveProduct = record
    Ean: Int64;
    Product: String;
    TypeProduct: String;
    Total: Int64;
    Amount: Float64;
    RecurringValue: Float64;
    ProvidersRecurringValue: String;

  end;

  TPricesCount = record
    Price: Float64;
    Qtde: Integer;
  end;

  TFormTopProdutos = class(TForm)
    Label1: TLabel;
    Image1: TImage;
    Panel1: TPanel;
    btnCSVImport: TBitBtn;
    btnExportXLSX: TBitBtn;
    sGrid: TStringGrid;
    pBar: TProgressBar;
    txtSearch: TEdit;
    lblCount: TLabel;
    txtProgress: TLabel;
    btnSearch: TButton;
    procedure FormCreate(Sender: TObject);
    procedure btnCSVImportClick(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
    procedure btnSearchClick(Sender: TObject);
    procedure txtSearchKeyUp(Sender: TObject; var Key: Word;
      Shift: TShiftState);
    procedure btnExportXLSXClick(Sender: TObject);
    procedure sGridSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
  private
    { private declarations }
  public

    CsvRecordsList: TList;
    ProductsList: TList;

    procedure LoadCSV(const fileSelected: String);
    procedure SeparateProducts;
    procedure GroupProducts;
    procedure PrintProducts;
    procedure PopulateFilteredList(const strSearch: String);
    procedure AutoSizeGridColumns(Grid: TStringGrid);
  end;

var
  FormTopProdutos: TFormTopProdutos;

implementation

uses uProducts, uCsv, ComObj;

{$R *.dfm}

function SortByAmount(A, B: Pointer): Integer;
var
  productOne, productTwo: TProduct;
begin

  productOne := TProduct(A);
  productTwo := TProduct(B);

  if productOne.Amount < productTwo.Amount then
    Result := 1
  else if productOne.Amount = productTwo.Amount then
    Result := 0
  else
    Result := -1;
end;

procedure TFormTopProdutos.btnCSVImportClick(Sender: TObject);
var
  dialog: TOpenDialog;
begin

  dialog := TOpenDialog.Create(nil);

  try
    dialog.Filter := 'Arquivo CSV (*.csv)| *.csv';

    if dialog.Execute(Handle) then
    begin

      CsvRecordsList := TList.Create;

      txtProgress.Caption := 'Lendo o arquivo...';
      txtProgress.Repaint;
      LoadCSV(dialog.FileName);

      txtProgress.Caption := 'Separando os produtos...';
      txtProgress.Repaint;
      SeparateProducts;

      txtProgress.Caption := 'Agrupando os produtos...';
      txtProgress.Repaint;
      GroupProducts;

      txtProgress.Caption := 'Montando os produtos na tela...';
      txtProgress.Repaint;
      PrintProducts;

      if ProductsList.Count > 0 then
      begin
        txtSearch.Enabled := True;
        btnSearch.Enabled := True;
        btnExportXLSX.Enabled := True;
      end
      else
      begin
        txtSearch.Enabled := False;
        btnSearch.Enabled := False;
        btnExportXLSX.Enabled := False;
      end;

      lblCount.Caption := 'Total de Produtos: ' + ProductsList.Count.ToString;
      txtProgress.Caption := 'Processo concluído.';
      txtProgress.Repaint;
      pBar.Position := 0;
    end;

  finally
    dialog.Free;
  end;
end;

procedure TFormTopProdutos.btnExportXLSXClick(Sender: TObject);
var
  Excel: Variant;
  Line: Int64;
  Product: TProduct;
begin

  Excel := CreateOleObject('Excel.Application');
  Excel.WorkBooks.add;
  Excel.Workbooks[1].WorkSheets[1].Name := 'TOPPRODUTOS';

  Excel.Workbooks[1].WorkSheets[1].cells[1,1].Value := 'Ranking de Produtos';
  Excel.Workbooks[1].WorkSheets[1].cells[1,1].Font.Name := 'Cambria';
  Excel.Workbooks[1].WorkSheets[1].cells[1,1].Font.Bold := True;
  Excel.Workbooks[1].WorkSheets[1].cells[1,1].Font.Size := 25;
  Excel.Workbooks[1].WorkSheets[1].Range['A1', 'G1'].MergeCells := True;

  {A} Excel.Workbooks[1].WorkSheets[1].cells[3,1].Value := 'EAN';
  {B} Excel.Workbooks[1].WorkSheets[1].cells[3,2].Value := 'PRODUTO';
  {C} Excel.Workbooks[1].WorkSheets[1].cells[3,3].Value := 'TIPO';
  {D} Excel.Workbooks[1].WorkSheets[1].cells[3,4].Value := 'QTDE';
  {E} Excel.Workbooks[1].WorkSheets[1].cells[3,5].Value := 'VALOR TOTAL COMPRADO';
  {G} Excel.Workbooks[1].WorkSheets[1].cells[3,6].Value := 'PREÇO RECORRENTE';
  {G} Excel.Workbooks[1].WorkSheets[1].cells[3,7].Value := 'FORNECEDOR (PREÇO RECORRENTE)';

  Excel.Workbooks[1].WorkSheets[1].Range['A3', 'G3'].Font.Name := 'Cambria';
  Excel.Workbooks[1].WorkSheets[1].Range['A3', 'G3'].Font.Size := 14;
  Excel.Workbooks[1].WorkSheets[1].Range['A3', 'G3'].Font.Bold := True;
  Excel.Workbooks[1].WorkSheets[1].Range['A3', 'A3'].HorizontalAlignment := 3;
  Excel.Workbooks[1].WorkSheets[1].Range['C3', 'F3'].HorizontalAlignment := 3;

  txtProgress.Caption := 'Gerando arquivo Excel...';
  txtProgress.Repaint;
  pBar.Max := ProductsList.Count;
  pBar.Position := 0;

  Line := 4;

  for Product in ProductsList do
  begin

    {A} Excel.Workbooks[1].WorkSheets[1].cells[Line,1].Value := Product.Ean;
    {B} Excel.Workbooks[1].WorkSheets[1].cells[Line,2].Value := Product.Product;
    {C} Excel.Workbooks[1].WorkSheets[1].cells[Line,3].Value := Product.Category;
    {D} Excel.Workbooks[1].WorkSheets[1].cells[Line,4].Value := Product.Count;
    {E} Excel.Workbooks[1].WorkSheets[1].cells[Line,5].Value := Product.Amount;
    {F} Excel.Workbooks[1].WorkSheets[1].cells[Line,6].Value := Product.RecurringPrice;
    {G} Excel.Workbooks[1].WorkSheets[1].cells[Line,7].Value := Product.Provider;

    Inc(Line);
    pBar.Position := pBar.Position +1;
  end;

  Excel.Workbooks[1].WorkSheets[1].Range['A4', ('G'+Line.ToString)].Font.Name := 'Cambria';
  Excel.Workbooks[1].WorkSheets[1].Range['A4', ('G'+Line.ToString)].Font.Size := 12;
  Excel.Workbooks[1].WorkSheets[1].Range['A4', ('A'+Line.ToString)].NumberFormat := '0000000000000';
  Excel.Workbooks[1].WorkSheets[1].Range['C4', ('F'+Line.ToString)].HorizontalAlignment := 3;
  Excel.Workbooks[1].WorkSheets[1].Range['E4', ('F'+Line.ToString)].NumberFormat := '_-R$ * #.##0,00_-;-R$ * #.##0,00_-;_-R$ * "-"??_-;_-@_-';

  Excel.Workbooks[1].WorkSheets[1].Columns.Autofit;

  Excel.Visible := True;

  txtProgress.Caption := 'Processo concluído.';
  txtProgress.Repaint;
  pBar.Position := 0;

end;

procedure TFormTopProdutos.btnSearchClick(Sender: TObject);
begin

  if Length(txtSearch.Text) > 0 then
  begin
    txtProgress.Caption := 'Pesquisando...';
    PopulateFilteredList(txtSearch.Text);
  end
  else
  begin
    PrintProducts;
  end;

  lblCount.Caption := 'Total de Produtos: ' + (sGrid.RowCount - 1).ToString;
  lblCount.Repaint;

  txtProgress.Caption := 'Processo concluído.';
  pBar.Position := 0;
end;

procedure TFormTopProdutos.LoadCSV(const fileSelected: String);
var
  Rows: TStrings;
  Cols: TStrings;
  Loop: Int64;
  CsvRecord: TCSVRecords;
begin

  Rows := TStringList.Create;
  Cols := TStringList.Create;

  Rows.Clear;
  Rows.LoadFromFile(fileSelected);

  CsvRecord := TCSVRecords.Create;

  if CsvRecord.CheckRowIntegrity(Rows) then
  begin

    pBar.Max := Rows.Count;

    for Loop := 1 to Rows.Count - 1 do
    begin

      Cols.Delimiter := ';';
      Cols.StrictDelimiter := True;
      Cols.DelimitedText := Rows[Loop];

      CsvRecord := TCSVRecords.Create;

      CsvRecord.Ean := Cols[0].ToInt64;
      CsvRecord.Product := Cols[1];
      CsvRecord.Category := Cols[2];
      CsvRecord.Provider := Cols[3];
      CsvRecord.Invoice := Cols[4].ToInt64;
      CsvRecord.Count := Cols[5].Replace('.', '').ToInt64;
      CsvRecord.Amount := Cols[6].Replace('.', '').ToDouble;
      CsvRecord.SetPrice;

      CsvRecordsList.Add(CsvRecord);

      pBar.Position := Loop;
    end;
  end
  else
  begin
    ShowMessage('Você escolheu o arquivo correto?');
  end;

end;

procedure TFormTopProdutos.SeparateProducts;
var
  CsvRecord: TCSVRecords;
  Product: TProduct;
begin

  ProductsList := TList.Create;

  pBar.Max := CsvRecordsList.Count;
  pBar.Position := 0;

  for CsvRecord in CsvRecordsList do
  begin

    Product := TProduct.Create;
    Product := Product.SearchObject(CsvRecord.Ean, ProductsList);

    if Product.Ean <> CsvRecord.Ean then
    begin

      Product.Ean := CsvRecord.Ean;
      Product.Product := CsvRecord.Product;
      Product.Category := CsvRecord.Category;
      Product.Count := CsvRecord.Count;
      Product.Amount := 0.0;
      Product.RecurringPrice := 0.0;
      Product.Provider := '';

      ProductsList.Add(Product);

    end;

    pBar.Position := pBar.Position + 1;
  end;
end;

procedure TFormTopProdutos.sGridSelectCell(Sender: TObject; ACol, ARow: Integer; var CanSelect: Boolean);
begin
  ShowMessage(sGrid.Cells[0, ARow]);
end;

procedure TFormTopProdutos.txtSearchKeyUp(Sender: TObject; var Key: Word;
  Shift: TShiftState);
begin

  if Key = 13 then
    btnSearchClick(Sender);

  if Length(txtSearch.Text) = 0 then
    if sGrid.RowCount < ProductsList.Count then
      PrintProducts;
      pBar.Position := 0;
end;

procedure TFormTopProdutos.GroupProducts;
var
  CsvRecord: TCSVRecords;
  Product, newProduct: TProduct;
  newProductsList: TList;
  prices: TArray<Double>;
begin

  newProductsList := TList.Create;

  pBar.Max := ProductsList.Count;
  pBar.Position := 0;

  for Product in ProductsList do
  begin

    newProduct := TProduct.Create;
    newProduct.Ean := Product.Ean;
    newProduct.Product := Product.Product;
    newProduct.Category := Product.Category;

    SetLength(prices, 0);

    for CsvRecord in CsvRecordsList do
    begin
      if newProduct.Ean = CsvRecord.Ean then
      begin
        newProduct.Count := newProduct.Count + CsvRecord.Count;
        newProduct.Amount := newProduct.Amount + CsvRecord.Amount;

        SetLength(prices, Length(prices) + 1);
        prices[Length(prices) - 1] := CsvRecord.Price;
      end;
    end;

    newProduct.RecurringPrice := Product.GetRecurringPrice(prices);

    for CsvRecord in CsvRecordsList do
    begin

      if CsvRecord.Price = newProduct.RecurringPrice then
      begin
        newProduct.Provider := CsvRecord.Provider;
        Break;
      end;

    end;

    newProductsList.Add(newProduct);

    pBar.Position := pBar.Position + 1;
  end;

  ProductsList := newProductsList;
  ProductsList.Sort(SortByAmount);
end;

procedure TFormTopProdutos.PrintProducts;
var
  Product: TProduct;
  Loop: Int64;
begin

  sGrid.RowCount := ProductsList.Count;
  sGrid.ColCount := 7;

  sGrid.Cells[0, 0] := 'EAN';
  sGrid.Cells[1, 0] := 'PRODUTO';
  sGrid.Cells[2, 0] := 'TIPO';
  sGrid.Cells[3, 0] := 'QTDE';
  sGrid.Cells[4, 0] := 'TOTAL COMPRADO';
  sGrid.Cells[5, 0] := 'PREÇO RECORRENTE';
  sGrid.Cells[6, 0] := 'FORNECEDOR (PREÇO RECORRENTE)';

  pBar.Max := ProductsList.Count;
  pBar.Position := 0;

  Loop := 1;

  for Product in ProductsList do
  begin
    sGrid.Cells[0, Loop] := Product.Ean.ToString;
    sGrid.Cells[1, Loop] := Product.Product;
    sGrid.Cells[2, Loop] := Product.Category;
    sGrid.Cells[3, Loop] := Product.Count.ToString;
    sGrid.Cells[4, Loop] := FormatFloat('R$ ###,###,#0.00', Product.Amount);
    sGrid.Cells[5, Loop] := FormatFloat('R$ ###,###,#0.00', Product.RecurringPrice);
    sGrid.Cells[6, Loop] := Product.Provider;

    pBar.Position := Loop;
    Inc(Loop);
  end;

  AutoSizeGridColumns(sGrid);
end;

procedure TFormTopProdutos.FormCreate(Sender: TObject);
begin

  Left := (Screen.Width - Width) div 2;
  Top := (Screen.Height - Height) div 2;
  KeyPreview := True;

  sGrid.ColCount := 7;

  sGrid.Cells[0, 0] := 'EAN';
  sGrid.Cells[1, 0] := 'PRODUTO';
  sGrid.Cells[2, 0] := 'TIPO';
  sGrid.Cells[3, 0] := 'QTDE';
  sGrid.Cells[4, 0] := 'TOTAL COMPRADO';
  sGrid.Cells[5, 0] := 'PREÇO RECORRENTE';
  sGrid.Cells[6, 0] := 'FORNECEDOR (PREÇO RECORRENTE)';

  AutoSizeGridColumns(sGrid);
end;

procedure TFormTopProdutos.PopulateFilteredList(const strSearch: string);
var
  RowCount, ColCount, NewRow: Int64;
  Found: Boolean;
  Product: TProduct;
begin

  if Length(strSearch) > 0 then
  begin

    RowCount := ProductsList.Count;
    ColCount := 7;

    sGrid.RowCount := 0;
    sGrid.ColCount := ColCount;

    sGrid.Cells[0, 0] := 'EAN';
    sGrid.Cells[1, 0] := 'PRODUTO';
    sGrid.Cells[2, 0] := 'TIPO';
    sGrid.Cells[3, 0] := 'QTDE';
    sGrid.Cells[4, 0] := 'TOTAL COMPRADO';
    sGrid.Cells[5, 0] := 'PREÇO RECORRENTE';
    sGrid.Cells[6, 0] := 'FORNECEDOR (PREÇO RECORRENTE)';

    NewRow := 1;

    pBar.Max := RowCount;
    pBar.Position := 0;

    for Product in ProductsList do
    begin

      Found := False;

      if Pos(LowerCase(strSearch), Product.Ean.ToString) <> 0 then Found := True;
      if Pos(LowerCase(strSearch), LowerCase(Product.Product)) <> 0 then Found := True;
      if Pos(LowerCase(strSearch), LowerCase(Product.Category)) <> 0 then Found := True;
      if Pos(LowerCase(strSearch), Product.Count.ToString) <> 0 then Found := True;
      if Pos(LowerCase(strSearch), FormatFloat('R$ ###,###,#0.00', Product.Amount)) <> 0 then Found := True;
      if Pos(LowerCase(strSearch), FormatFloat('R$ ###,###,#0.00', Product.RecurringPrice)) <> 0 then Found := True;
      if Pos(LowerCase(strSearch), LowerCase(Product.Provider)) <> 0 then Found := True;

      if Found = True then
      begin
        sGrid.RowCount := NewRow + 1;
        sGrid.Cells[0, NewRow] := Product.Ean.ToString;
        sGrid.Cells[1, NewRow] := Product.Product;
        sGrid.Cells[2, NewRow] := Product.Category;
        sGrid.Cells[3, NewRow] := Product.Count.ToString;
        sGrid.Cells[4, NewRow] := FormatFloat('R$ ###,###,#0.00', Product.Amount);
        sGrid.Cells[5, NewRow] := FormatFloat('R$ ###,###,#0.00', Product.RecurringPrice);
        sGrid.Cells[6, NewRow] := Product.Provider;

        NewRow := NewRow + 1;
      end;

      pBar.Position := pBar.Position + 1;

    end;
  end;

  if (sGrid.RowCount - 1) = 0 then
  begin
    PrintProducts;
    ShowMessage('A sua pesquisa não corresponde a nada nesta lista');
  end;


  lblCount.Caption := 'Total de Produtos: ' + (sGrid.RowCount - 1).ToString;
  lblCount.Repaint;

end;

procedure TFormTopProdutos.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #27 then
    Close;
end;

procedure TFormTopProdutos.AutoSizeGridColumns(Grid: TStringGrid);
const
  MIN_COL_WIDTH = 15;
var
  Col: Integer;
  ColWidth, CellWidth: Integer;
  Row: Integer;
begin
  Grid.Canvas.Font.Assign(Grid.Font);
  for Col := 0 to Grid.ColCount - 1 do
  begin
    ColWidth := Grid.Canvas.TextWidth(Grid.Cells[Col, 0]);
    for Row := 0 to Grid.RowCount - 1 do
    begin
      CellWidth := Grid.Canvas.TextWidth(Grid.Cells[Col, Row]);
      if CellWidth > ColWidth then
        ColWidth := CellWidth
    end;
    Grid.ColWidths[Col] := ColWidth + MIN_COL_WIDTH;
  end;
end;

end.
