unit uImportarProdutos;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, DateTimePicker, uImportarMarcas, uImportarGrupos,
  uImportarTributacao, uAnalisadorVariantes, uImportadorBase;

type

  { TForm2 }

  TForm2 = class(TForm)
    ButtonLimparErros: TButton;
    ButtonAbrirTributacao: TButton;
    ButtonAbrirGrupos: TButton;
    ButtonAbrirMarcas: TButton;
    ButtonImportar: TButton;
    ButtonBuscarFDB: TButton;
    ButtonBuscarDBF: TButton;
    CheckUsarCodigoAlternativo: TCheckBox;
    CheckUsarGrade: TCheckBox;
    CheckEstoque: TCheckBox;
    ComboSituTributa: TComboBox;
    ComboCodigoAlternativo: TComboBox;
    ComboClassificacaoFiscal: TComboBox;
    ComboEstoque: TComboBox;
    ComboPrecoEntrada: TComboBox;
    ComboPrecoCusto: TComboBox;
    ComboPrecoVenda: TComboBox;
    ComboMarca: TComboBox;
    ComboGrupo: TComboBox;
    ComboNCM: TComboBox;
    ComboUnidade: TComboBox;
    ComboCodBarras: TComboBox;
    ComboDescricao: TComboBox;
    EditICMS: TEdit;
    EditLinhaInicial: TEdit;
    EditCaminhoFDB: TEdit;
    EditCaminhoDBF: TEdit;
    LabelICMS: TLabel;
    LabelErros: TLabel;
    LabelSituTributa: TLabel;
    LabelProgresso: TLabel;
    LabelCodigoAlternativo: TLabel;
    LabelClassificacaoFiscal: TLabel;
    LabelEstoque: TLabel;
    LabelPrecoEntrada: TLabel;
    LabelPrecoCusto: TLabel;
    LabelPrecoVenda: TLabel;
    LabelMarca: TLabel;
    LabelGrupo: TLabel;
    LabelNCM: TLabel;
    LabelUnidade: TLabel;
    LabelCodBarras: TLabel;
    LabelDescricao: TLabel;
    LabelLinhaInicial: TLabel;
    LabelFDB: TLabel;
    LabelDBF: TLabel;
    MemoErros: TMemo;
    OpenDialog: TOpenDialog;
    ProgressBar: TProgressBar;

    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarClick(Sender: TObject);
    procedure ComboSituTributaChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure ButtonAbrirMarcasClick(Sender: TObject);
    procedure ButtonAbrirGruposClick(Sender: TObject);
    procedure ButtonAbrirTributacaoClick(Sender: TObject);
    procedure ButtonLimparErrosClick(Sender: TObject);
    procedure CheckUsarCodigoAlternativoClick(Sender: TObject);
    procedure LabelSituTributaClick(Sender: TObject);

  private
    FImportadorBase: TImportadorBase;
    FPortaFirebird: Integer;

    procedure CarregarColunas;
    procedure ImportarProdutos;
    procedure ImportarSomenteGrupos;
    procedure ImportarSomenteMarcas;
    procedure ImportarSomenteTributacao;
    procedure ImportarProdutosCompleto;
    procedure AtualizarProgresso(APosicaoAtual: Integer);
    procedure AtualizarLabelProgresso(const ATextoProgresso: string);
    procedure LogErro(const ALinha: Integer; const AMensagemErro: string; const ADescricao: string);

  public
    property PortaFirebird: Integer read FPortaFirebird write FPortaFirebird;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  Form2: TForm2;

implementation

{$R *.lfm}

{ TForm2 }

constructor TForm2.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
  FImportadorBase := TImportadorBase.Create;
  FImportadorBase.OnProgress := @AtualizarProgresso;
  FImportadorBase.OnProgressLabel := @AtualizarLabelProgresso;
  FImportadorBase.OnLogErro := @LogErro;
end;

destructor TForm2.Destroy;
begin
  FImportadorBase.Free;
  inherited Destroy;
end;

procedure TForm2.AtualizarProgresso(APosicaoAtual: Integer);
begin
  ProgressBar.Position := APosicaoAtual;
  Application.ProcessMessages;
end;

procedure TForm2.AtualizarLabelProgresso(const ATextoProgresso: string);
begin
  LabelProgresso.Caption := ATextoProgresso;
  Application.ProcessMessages;
end;

procedure TForm2.LogErro(const ALinha: Integer; const AMensagemErro: string; const ADescricao: string);
begin
  MemoErros.Lines.Add(Format('Linha %d: %s - %s', [ALinha, AMensagemErro, ADescricao]));
  Application.ProcessMessages;
end;

procedure TForm2.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor - Produtos';
  Width := 1100;
  Height := 800;
  Position := poScreenCenter;
end;

procedure TForm2.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
end;

procedure TForm2.CheckUsarCodigoAlternativoClick(Sender: TObject);
begin
  ComboCodigoAlternativo.Enabled := CheckUsarCodigoAlternativo.Checked;
end;

procedure TForm2.LabelSituTributaClick(Sender: TObject);
begin

end;

procedure TForm2.ButtonAbrirMarcasClick(Sender: TObject);
begin
  if not Assigned(Form6) then
    Form6 := TForm6.Create(Application);
  Form6.Show;
end;

procedure TForm2.ButtonAbrirGruposClick(Sender: TObject);
begin
  if not Assigned(Form7) then
    Form7 := TForm7.Create(Application);
  Form7.Show;
end;

procedure TForm2.ButtonAbrirTributacaoClick(Sender: TObject);
begin
  if not Assigned(Form8) then
    Form8 := TForm8.Create(Application);
  Form8.Show;
end;

procedure TForm2.ButtonLimparErrosClick(Sender: TObject);
begin
  MemoErros.Clear;
end;

procedure TForm2.ButtonBuscarDBFClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Arquivos DBF|*.dbf';
  if OpenDialog.Execute then
  begin
    EditCaminhoDBF.Text := OpenDialog.FileName;
    CarregarColunas;
  end;
end;

procedure TForm2.ButtonBuscarFDBClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Banco Firebird|*.fdb';
  if OpenDialog.Execute then
    EditCaminhoFDB.Text := OpenDialog.FileName;
end;

procedure TForm2.ButtonImportarClick(Sender: TObject);
begin
  if (EditCaminhoDBF.Text = '') or (EditCaminhoFDB.Text = '') then
  begin
    ShowMessage('Selecione os arquivos DBF e FDB antes de importar.');
    Exit;
  end;

  MemoErros.Clear;
  FImportadorBase.LinhasComErro.Clear;

  FImportadorBase.Mapeamento.Values['DESCRICAO'] := ComboDescricao.Text;
  FImportadorBase.Mapeamento.Values['BARRAS'] := ComboCodBarras.Text;
  FImportadorBase.Mapeamento.Values['UN'] := ComboUnidade.Text;
  FImportadorBase.Mapeamento.Values['NCM'] := ComboNCM.Text;
  FImportadorBase.Mapeamento.Values['PRECO_VENDA'] := ComboPrecoVenda.Text;
  FImportadorBase.Mapeamento.Values['PRECO_CUSTO'] := ComboPrecoCusto.Text;
  FImportadorBase.Mapeamento.Values['PRECO_ENTRADA'] := ComboPrecoEntrada.Text;
  FImportadorBase.Mapeamento.Values['ESTOQUE'] := ComboEstoque.Text;
  FImportadorBase.Mapeamento.Values['GRUPO'] := ComboGrupo.Text;
  FImportadorBase.Mapeamento.Values['MARCA'] := ComboMarca.Text;
  FImportadorBase.Mapeamento.Values['CLASSIFICACAO_FISCAL'] := ComboClassificacaoFiscal.Text;

  FImportadorBase.UsarCodigoAlternativo := CheckUsarCodigoAlternativo.Checked;
  FImportadorBase.CodigoAlternativoField := ComboCodigoAlternativo.Text;

  FImportadorBase.CaminhoDBF := EditCaminhoDBF.Text;
  FImportadorBase.CaminhoFDB := EditCaminhoFDB.Text;
  FImportadorBase.ImportarEstoque := CheckEstoque.Checked;
  FImportadorBase.UsarGrade := CheckUsarGrade.Checked;
  FImportadorBase.LinhaInicial := StrToIntDef(EditLinhaInicial.Text, 1);
  FImportadorBase.SituTributaria := Copy(ComboSituTributa.Text, 1, 1);
  FImportadorBase.EditICMS := EditICMS.Text;

  try
    ImportarProdutosCompleto;

    if FImportadorBase.LinhasComErro.Count > 0 then
      ShowMessage(Format('Importação concluída com %d erros. Verifique a lista de erros.', [FImportadorBase.LinhasComErro.Count]))
    else
      ShowMessage('Importação concluída com sucesso!');

  except
    on E: Exception do
      ShowMessage('Erro durante a importação: ' + E.Message);
  end;
end;

procedure TForm2.ComboSituTributaChange(Sender: TObject);
begin
  if (ComboSituTributa.ItemIndex = 1) or (ComboSituTributa.ItemIndex = 2) or (ComboSituTributa.ItemIndex = 3) or (ComboSituTributa.ItemIndex = 5) then
     EditICMS.Enabled:=False
  else
     EditICMS.Enabled := True;

end;

procedure TForm2.CarregarColunas;
var
  dbfProdutos: TDbf;
  i: Integer;
begin
  if not FileExists(EditCaminhoDBF.Text) then
    Exit;

  dbfProdutos := TDbf.Create(nil);
  try
    dbfProdutos.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
    dbfProdutos.TableName := ExtractFileName(EditCaminhoDBF.Text);
    dbfProdutos.Open;

    ComboDescricao.Clear;
    ComboCodBarras.Clear;
    ComboUnidade.Clear;
    ComboNCM.Clear;
    ComboPrecoVenda.Clear;
    ComboPrecoCusto.Clear;
    ComboPrecoEntrada.Clear;
    ComboEstoque.Clear;
    ComboGrupo.Clear;
    ComboMarca.Clear;
    ComboClassificacaoFiscal.Clear;
    ComboCodigoAlternativo.Clear;

    for i := 0 to dbfProdutos.FieldDefs.Count - 1 do
    begin
      ComboDescricao.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboCodBarras.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboUnidade.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboNCM.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboPrecoVenda.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboPrecoCusto.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboPrecoEntrada.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboEstoque.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboGrupo.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboMarca.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboClassificacaoFiscal.Items.Add(dbfProdutos.FieldDefs[i].Name);
      ComboCodigoAlternativo.Items.Add(dbfProdutos.FieldDefs[i].Name);
    end;

    dbfProdutos.Close;
    ShowMessage('Colunas carregadas com sucesso!');
  finally
    dbfProdutos.Free;
  end;
end;

procedure TForm2.ImportarProdutos;
begin
  FImportadorBase.ImportarProdutos;
end;

procedure TForm2.ImportarSomenteGrupos;
begin
  FImportadorBase.ImportarSomenteGrupos;
end;

procedure TForm2.ImportarSomenteMarcas;
begin
  FImportadorBase.ImportarSomenteMarcas;
end;

procedure TForm2.ImportarSomenteTributacao;
begin
  FImportadorBase.ImportarSomenteTributacao;
end;

procedure TForm2.ImportarProdutosCompleto;
begin
  ImportarProdutos;
end;

end.
