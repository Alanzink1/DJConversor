unit uPrincipal;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, Forms, Controls, Graphics, Dialogs, StdCtrls;

type
  { TForm1 }
  TForm1 = class(TForm)
    ButtonProdutos: TButton;
    ButtonClientes: TButton;
    ButtonContasReceber: TButton;
    LabelTitulo: TLabel;
    ButtonConfiguracao: TButton;

    procedure ButtonProdutosClick(Sender: TObject);
    procedure ButtonClientesClick(Sender: TObject);
    procedure ButtonContasReceberClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure ButtonConfiguracaoClick(Sender: TObject);

  private
    FPortaFirebird: Integer;
    procedure SalvarConfiguracao;
    procedure CarregarConfiguracao;

  public
    property PortaFirebird: Integer read FPortaFirebird write FPortaFirebird;
  end;

var
  Form1: TForm1;

implementation

uses
  uImportarProdutos, uImportarClientes, uImportarContas;

{$R *.lfm}

{ TFPrincipal }

procedure TForm1.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor';
  Width := 400;
  Height := 330;

  CarregarConfiguracao;

  LabelTitulo := TLabel.Create(Self);
  LabelTitulo.Parent := Self;
  LabelTitulo.Caption := 'SISTEMA DE IMPORTAÇÃO';
  LabelTitulo.Font.Size := 16;
  LabelTitulo.Font.Style := [fsBold];
  LabelTitulo.Left := 75;
  LabelTitulo.Top := 40;

  ButtonConfiguracao := TButton.Create(Self);
  ButtonConfiguracao.Parent := Self;
  ButtonConfiguracao.Caption := '⚙️ Configuração';
  ButtonConfiguracao.Left := 100;
  ButtonConfiguracao.Top := 240;
  ButtonConfiguracao.Width := 200;
  ButtonConfiguracao.Height := 40;
  ButtonConfiguracao.OnClick := @ButtonConfiguracaoClick;

  ButtonProdutos := TButton.Create(Self);
  ButtonProdutos.Parent := Self;
  ButtonProdutos.Caption := '📦 Importar Produtos';
  ButtonProdutos.Left := 100;
  ButtonProdutos.Top := 90;
  ButtonProdutos.Width := 200;
  ButtonProdutos.Height := 40;
  ButtonProdutos.OnClick := @ButtonProdutosClick;

  ButtonClientes := TButton.Create(Self);
  ButtonClientes.Parent := Self;
  ButtonClientes.Caption := '👥 Importar Clientes';
  ButtonClientes.Left := 100;
  ButtonClientes.Top := 140;
  ButtonClientes.Width := 200;
  ButtonClientes.Height := 40;
  ButtonClientes.OnClick := @ButtonClientesClick;

  ButtonContasReceber := TButton.Create(Self);
  ButtonContasReceber.Parent := Self;
  ButtonContasReceber.Caption := '💰 Importar Contas a Receber';
  ButtonContasReceber.Left := 100;
  ButtonContasReceber.Top := 190;
  ButtonContasReceber.Width := 200;
  ButtonContasReceber.Height := 40;
  ButtonContasReceber.OnClick := @ButtonContasReceberClick;
end;


procedure TForm1.ButtonConfiguracaoClick(Sender: TObject);
var
  PortaStr: string;
  NovaPorta: Integer;
begin

  PortaStr := IntToStr(FPortaFirebird);
  if InputQuery('Configuração do Firebird', sLineBreak +
                'Digite a porta do Firebird:', PortaStr) then
  begin
    if TryStrToInt(PortaStr, NovaPorta) and (NovaPorta > 0) then
    begin
      FPortaFirebird := NovaPorta;
      SalvarConfiguracao;
      ShowMessage('Porta configurada para: ' + IntToStr(FPortaFirebird));
    end
    else
    begin
      ShowMessage('Porta inválida! Usando porta padrão 3050.');
      FPortaFirebird := 3050;
    end;
  end;
end;


procedure TForm1.SalvarConfiguracao;
var
  ConfigFile: TStringList;
begin
  ConfigFile := TStringList.Create;
  try
    ConfigFile.Values['PortaFirebird'] := IntToStr(FPortaFirebird);
    ConfigFile.SaveToFile(ExtractFilePath(ParamStr(0)) + 'config.ini');
  finally
    ConfigFile.Free;
  end;
end;


procedure TForm1.CarregarConfiguracao;
var
  ConfigFile: TStringList;
  ConfigPath: string;
begin
  ConfigPath := ExtractFilePath(ParamStr(0)) + 'config.ini';

  if FileExists(ConfigPath) then
  begin
    ConfigFile := TStringList.Create;
    try
      ConfigFile.LoadFromFile(ConfigPath);
      FPortaFirebird := StrToIntDef(ConfigFile.Values['PortaFirebird'], 3050);
    finally
      ConfigFile.Free;
    end;
  end
  else
  begin

    FPortaFirebird := 3050;
  end;
end;

procedure TForm1.ButtonProdutosClick(Sender: TObject);
begin
  Form2 := TForm2.Create(Application);
  try

    Form2.PortaFirebird := Self.PortaFirebird;
    Form2.ShowModal;
  finally
    Form2.Free;
  end;
end;

procedure TForm1.ButtonClientesClick(Sender: TObject);
begin
  Form3 := TForm3.Create(Application);
  try

    Form3.PortaFirebird := Self.PortaFirebird;
    Form3.ShowModal;
  finally
    Form3.Free;
  end;
end;

procedure TForm1.ButtonContasReceberClick(Sender: TObject);
begin
  Form4 := TForm4.Create(Application);
  try

    Form4.PortaFirebird := Self.PortaFirebird;
    Form4.ShowModal;
  finally
    Form4.Free;
  end;
end;

end.
