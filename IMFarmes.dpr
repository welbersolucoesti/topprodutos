program IMFarmes;

uses
  Vcl.Forms,
  Home in 'Home.pas' {FormHome},
  TopProdutos in 'TopProdutos.pas' {FormTopProdutos},
  Vcl.Themes,
  Vcl.Styles,
  uCsv in 'uCsv.pas';

{$R *.res}

begin
  Application.Initialize;
  Application.MainFormOnTaskbar := True;
  TStyleManager.TrySetStyle('Windows10 Charcoal');
  Application.CreateForm(TFormHome, FormHome);
  Application.CreateForm(TFormTopProdutos, FormTopProdutos);
  Application.Run;
end.
