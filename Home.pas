unit Home;

interface

uses
  Winapi.Windows, Winapi.Messages, System.SysUtils, System.Variants,
  System.Classes, Vcl.Graphics,
  Vcl.Controls, Vcl.Forms, Vcl.Dialogs, Vcl.ExtCtrls, Vcl.Menus,
  Vcl.Imaging.pngimage, Vcl.StdCtrls, Vcl.Buttons;

type
  TFormHome = class(TForm)
    head: TPanel;
    MainMenu1: TMainMenu;
    Aplicativos1: TMenuItem;
    opProdutos1: TMenuItem;
    Sistema1: TMenuItem;
    Sair1: TMenuItem;
    BitBtn1: TBitBtn;
    procedure opProdutos1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure Sair1Click(Sender: TObject);
    procedure BitBtn1Click(Sender: TObject);
    procedure FormKeyPress(Sender: TObject; var Key: Char);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  FormHome: TFormHome;

implementation

{$R *.dfm}

uses TopProdutos;

procedure TFormHome.BitBtn1Click(Sender: TObject);
begin
  FormTopProdutos.ShowModal;
end;

procedure TFormHome.FormCreate(Sender: TObject);
begin
  Left := (Screen.Width - Width) div 2;
  Top := (Screen.Height - Height) div 2;
  KeyPreview := True;
end;

procedure TFormHome.FormKeyPress(Sender: TObject; var Key: Char);
begin
  if Key = #27 then Close;
end;

procedure TFormHome.opProdutos1Click(Sender: TObject);
begin
  FormTopProdutos.ShowModal;
end;

procedure TFormHome.Sair1Click(Sender: TObject);
begin
  Application.Terminate;
end;

end.
