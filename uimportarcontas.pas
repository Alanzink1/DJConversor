unit uImportarContas;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, Variants, StrUtils, DateUtils;

type
  { TForm4 }
  TForm4 = class(TForm)
    ButtonBuscarDBF: TButton;
    ButtonBuscarFDB: TButton;
    ButtonImportar: TButton;
    EditCaminhoDBF: TEdit;
    EditCaminhoFDB: TEdit;
    LabelDBF: TLabel;
    LabelFDB: TLabel;
    LabelNome: TLabel;
    LabelDataCx: TLabel;
    LabelVenc: TLabel;
    LabelAtraso: TLabel;
    LabelValor: TLabel;
    LabelJuros: TLabel;
    LabelValorConta: TLabel;
    ComboNome: TComboBox;
    ComboDataCx: TComboBox;
    ComboVenc: TComboBox;
    ComboAtraso: TComboBox;
    ComboValor: TComboBox;
    ComboJuros: TComboBox;
    ComboValorConta: TComboBox;
    GroupParams: TGroupBox;
    LabelLoja: TLabel;
    LabelTerminal: TLabel;
    LabelNatureza: TLabel;
    LabelPlano: TLabel;
    LabelCarteira: TLabel;
    EditCodLoja: TEdit;
    EditCodTerminal: TEdit;
    EditIdNatureza: TEdit;
    EditCodPlanoPagto: TEdit;
    EditCodCarteira: TEdit;
    ProgressBar: TProgressBar;
    LabelProgresso: TLabel;
    OpenDialog: TOpenDialog;

    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);

  private
    FPortaFirebird: Integer;
    LastIDFatura: Integer;

    procedure CarregarColunas;
    procedure ImportarContasReceber;
    function ParseDateBR(const S: string): TDateTime;
    function ParseFloatBR(const S: string): Double;
    procedure WarmUpFaturaCounter(Q: TSQLQuery);
    function NextIDFatura(Q: TSQLQuery): Integer;
    procedure LogMsg(const nomeArq, msg: string);
    function RemoveAcentos(const S: string): string;
    function PrepareStringForDB(const S: string): string;
    procedure Prep(Q: TSQLQuery; const SQLText: string);

  public
    property PortaFirebird: Integer read FPortaFirebird write FPortaFirebird;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  Form4: TForm4;

implementation

{$R *.lfm}

const
  LOGS_ENABLED = True;

{ TForm4 }

constructor TForm4.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TForm4.Destroy;
begin
  inherited Destroy;
end;

procedure TForm4.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor - Contas a Receber';
  Width := 900;
  Height := 650;
  Position := poScreenCenter;


  LabelDBF := TLabel.Create(Self);
  LabelDBF.Parent := Self;
  LabelDBF.Caption := 'Arquivo DBF:';
  LabelDBF.Left := 20;
  LabelDBF.Top := 20;

  EditCaminhoDBF := TEdit.Create(Self);
  EditCaminhoDBF.Parent := Self;
  EditCaminhoDBF.Left := 150;
  EditCaminhoDBF.Top := 16;
  EditCaminhoDBF.Width := 500;

  ButtonBuscarDBF := TButton.Create(Self);
  ButtonBuscarDBF.Parent := Self;
  ButtonBuscarDBF.Caption := '📂 Buscar';
  ButtonBuscarDBF.Left := 670;
  ButtonBuscarDBF.Top := 16;
  ButtonBuscarDBF.Width := 80;
  ButtonBuscarDBF.OnClick := @ButtonBuscarDBFClick;

  LabelFDB := TLabel.Create(Self);
  LabelFDB.Parent := Self;
  LabelFDB.Caption := 'Banco FDB:';
  LabelFDB.Left := 20;
  LabelFDB.Top := 60;

  EditCaminhoFDB := TEdit.Create(Self);
  EditCaminhoFDB.Parent := Self;
  EditCaminhoFDB.Left := 150;
  EditCaminhoFDB.Top := 56;
  EditCaminhoFDB.Width := 500;

  ButtonBuscarFDB := TButton.Create(Self);
  ButtonBuscarFDB.Parent := Self;
  ButtonBuscarFDB.Caption := '📁 Buscar';
  ButtonBuscarFDB.Left := 670;
  ButtonBuscarFDB.Top := 56;
  ButtonBuscarFDB.Width := 80;
  ButtonBuscarFDB.OnClick := @ButtonBuscarFDBClick;


  LabelNome := TLabel.Create(Self);
  LabelNome.Parent := Self;
  LabelNome.Caption := 'NOME:';
  LabelNome.Left := 20;
  LabelNome.Top := 110;

  ComboNome := TComboBox.Create(Self);
  ComboNome.Parent := Self;
  ComboNome.Left := 150;
  ComboNome.Top := 106;
  ComboNome.Width := 200;

  LabelDataCx := TLabel.Create(Self);
  LabelDataCx.Parent := Self;
  LabelDataCx.Caption := 'DATACX:';
  LabelDataCx.Left := 20;
  LabelDataCx.Top := 150;

  ComboDataCx := TComboBox.Create(Self);
  ComboDataCx.Parent := Self;
  ComboDataCx.Left := 150;
  ComboDataCx.Top := 146;
  ComboDataCx.Width := 200;

  LabelVenc := TLabel.Create(Self);
  LabelVenc.Parent := Self;
  LabelVenc.Caption := 'VENC:';
  LabelVenc.Left := 20;
  LabelVenc.Top := 190;

  ComboVenc := TComboBox.Create(Self);
  ComboVenc.Parent := Self;
  ComboVenc.Left := 150;
  ComboVenc.Top := 186;
  ComboVenc.Width := 200;

  LabelAtraso := TLabel.Create(Self);
  LabelAtraso.Parent := Self;
  LabelAtraso.Caption := 'ATRASO:';
  LabelAtraso.Left := 20;
  LabelAtraso.Top := 230;

  ComboAtraso := TComboBox.Create(Self);
  ComboAtraso.Parent := Self;
  ComboAtraso.Left := 150;
  ComboAtraso.Top := 226;
  ComboAtraso.Width := 200;

  LabelValor := TLabel.Create(Self);
  LabelValor.Parent := Self;
  LabelValor.Caption := 'VALOR:';
  LabelValor.Left := 400;
  LabelValor.Top := 110;

  ComboValor := TComboBox.Create(Self);
  ComboValor.Parent := Self;
  ComboValor.Left := 500;
  ComboValor.Top := 106;
  ComboValor.Width := 200;

  LabelJuros := TLabel.Create(Self);
  LabelJuros.Parent := Self;
  LabelJuros.Caption := 'N_JUROS:';
  LabelJuros.Left := 400;
  LabelJuros.Top := 150;

  ComboJuros := TComboBox.Create(Self);
  ComboJuros.Parent := Self;
  ComboJuros.Left := 500;
  ComboJuros.Top := 146;
  ComboJuros.Width := 200;

  LabelValorConta := TLabel.Create(Self);
  LabelValorConta.Parent := Self;
  LabelValorConta.Caption := 'VALORCONTA:';
  LabelValorConta.Left := 400;
  LabelValorConta.Top := 190;

  ComboValorConta := TComboBox.Create(Self);
  ComboValorConta.Parent := Self;
  ComboValorConta.Left := 500;
  ComboValorConta.Top := 186;
  ComboValorConta.Width := 200;


  GroupParams := TGroupBox.Create(Self);
  GroupParams.Parent := Self;
  GroupParams.Caption := 'Parâmetros do Sistema';
  GroupParams.Left := 20;
  GroupParams.Top := 270;
  GroupParams.Width := 750;
  GroupParams.Height := 150;

  LabelLoja := TLabel.Create(Self);
  LabelLoja.Parent := GroupParams;
  LabelLoja.Caption := 'Código Loja:';
  LabelLoja.Left := 20;
  LabelLoja.Top := 25;

  EditCodLoja := TEdit.Create(Self);
  EditCodLoja.Parent := GroupParams;
  EditCodLoja.Left := 120;
  EditCodLoja.Top := 22;
  EditCodLoja.Width := 80;
  EditCodLoja.Text := '1';

  LabelTerminal := TLabel.Create(Self);
  LabelTerminal.Parent := GroupParams;
  LabelTerminal.Caption := 'Código Terminal:';
  LabelTerminal.Left := 220;
  LabelTerminal.Top := 25;

  EditCodTerminal := TEdit.Create(Self);
  EditCodTerminal.Parent := GroupParams;
  EditCodTerminal.Left := 340;
  EditCodTerminal.Top := 22;
  EditCodTerminal.Width := 80;
  EditCodTerminal.Text := '1';

  LabelNatureza := TLabel.Create(Self);
  LabelNatureza.Parent := GroupParams;
  LabelNatureza.Caption := 'ID Natureza:';
  LabelNatureza.Left := 440;
  LabelNatureza.Top := 25;

  EditIdNatureza := TEdit.Create(Self);
  EditIdNatureza.Parent := GroupParams;
  EditIdNatureza.Left := 540;
  EditIdNatureza.Top := 22;
  EditIdNatureza.Width := 80;

  LabelPlano := TLabel.Create(Self);
  LabelPlano.Parent := GroupParams;
  LabelPlano.Caption := 'Código Plano:';
  LabelPlano.Left := 20;
  LabelPlano.Top := 65;

  EditCodPlanoPagto := TEdit.Create(Self);
  EditCodPlanoPagto.Parent := GroupParams;
  EditCodPlanoPagto.Left := 120;
  EditCodPlanoPagto.Top := 62;
  EditCodPlanoPagto.Width := 80;
  EditCodPlanoPagto.Text := '1';

  LabelCarteira := TLabel.Create(Self);
  LabelCarteira.Parent := GroupParams;
  LabelCarteira.Caption := 'Código Carteira:';
  LabelCarteira.Left := 220;
  LabelCarteira.Top := 65;

  EditCodCarteira := TEdit.Create(Self);
  EditCodCarteira.Parent := GroupParams;
  EditCodCarteira.Left := 340;
  EditCodCarteira.Top := 62;
  EditCodCarteira.Width := 80;
  EditCodCarteira.Text := '1';

  ButtonImportar := TButton.Create(Self);
  ButtonImportar.Parent := Self;
  ButtonImportar.Caption := '🚀 Importar Contas a Receber';
  ButtonImportar.Left := 20;
  ButtonImportar.Top := 440;
  ButtonImportar.Width := 250;
  ButtonImportar.Height := 40;
  ButtonImportar.OnClick := @ButtonImportarClick;

  ProgressBar := TProgressBar.Create(Self);
  ProgressBar.Parent := Self;
  ProgressBar.Left := 20;
  ProgressBar.Top := 500;
  ProgressBar.Width := 750;
  ProgressBar.Min := 0;
  ProgressBar.Max := 100;

  LabelProgresso := TLabel.Create(Self);
  LabelProgresso.Parent := Self;
  LabelProgresso.Left := 20;
  LabelProgresso.Top := 530;
  LabelProgresso.Caption := '';

  OpenDialog := TOpenDialog.Create(Self);
end;

procedure TForm4.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
end;

{ MÉTODOS UTILITÁRIOS }

procedure TForm4.LogMsg(const nomeArq, msg: string);
var
  f: TextFile;
  caminho: string;
begin
  if not LOGS_ENABLED then Exit;

  caminho := ExtractFilePath(ParamStr(0)) + nomeArq;
  AssignFile(f, caminho);
  try
    if FileExists(caminho) then
      Append(f)
    else
      Rewrite(f);
    Writeln(f, FormatDateTime('yyyy-mm-dd hh:nn:ss', Now) + ' - ' + msg);
  finally
    CloseFile(f);
  end;
end;

function TForm4.RemoveAcentos(const S: string): string;
const
  ComAcento = 'ÀÁÂÃÄÅÇÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïòóôõöùúûüýÿ';
  SemAcento = 'AAAAAACEEEEIIIIOOOOOUUUUYaaaaaaceeeeiiiiooooouuuuyy';
var
  i, j: Integer;
begin
  Result := S;
  for i := 1 to Length(ComAcento) do
    for j := 1 to Length(Result) do
      if Result[j] = ComAcento[i] then
        Result[j] := SemAcento[i];
end;

function TForm4.PrepareStringForDB(const S: string): string;
var
  i: Integer;
  temp: string;
begin
  temp := RemoveAcentos(Trim(S));


  Result := '';
  for i := 1 to Length(temp) do
  begin
    if CharInSet(temp[i],
       ['A'..'Z', 'a'..'z', '0'..'9', ' ', '.', ',', '-', '_', '/', '(', ')',
        '[', ']', '{', '}', '!', '@', '#', '$', '%', '*', '+', '=', ':', ';',
        '<', '>', '?', '|', '~', '`', '&', '^']) then
    begin
      Result := Result + temp[i];
    end
    else if temp[i] = '''' then
    begin
      Result := Result + '"';
    end;
  end;

  Result := Trim(Result);
end;

procedure TForm4.Prep(Q: TSQLQuery; const SQLText: string);
begin
  Q.Close;
  if Q.Prepared then
    Q.UnPrepare;
  Q.SQL.Text := SQLText;
  Q.Prepare;
end;

{ MÉTODOS DE UI }

procedure TForm4.ButtonBuscarDBFClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Arquivos DBF|*.dbf';
  if OpenDialog.Execute then
  begin
    EditCaminhoDBF.Text := OpenDialog.FileName;
    CarregarColunas;
  end;
end;

procedure TForm4.ButtonBuscarFDBClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Banco Firebird|*.fdb';
  if OpenDialog.Execute then
    EditCaminhoFDB.Text := OpenDialog.FileName;
end;

procedure TForm4.ButtonImportarClick(Sender: TObject);
begin
  if (EditCaminhoDBF.Text = '') or (EditCaminhoFDB.Text = '') then
  begin
    ShowMessage('Selecione os arquivos DBF e FDB antes de importar.');
    Exit;
  end;


  if ComboNome.Text = '' then ComboNome.Text := 'NOME';
  if ComboDataCx.Text = '' then ComboDataCx.Text := 'DATACX';
  if ComboVenc.Text = '' then ComboVenc.Text := 'VENC';
  if ComboAtraso.Text = '' then ComboAtraso.Text := 'ATRASO';
  if ComboValor.Text = '' then ComboValor.Text := 'VALORCONTA';
  if ComboJuros.Text = '' then ComboJuros.Text := 'N_JUROS';
  if ComboValorConta.Text = '' then ComboValorConta.Text := 'VALORCONTA';

  ImportarContasReceber;
end;

procedure TForm4.CarregarColunas;
var
  dbc: TDbf;
  i: Integer;
begin
  if not FileExists(EditCaminhoDBF.Text) then
    Exit;

  dbc := TDbf.Create(nil);
  try
    dbc.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
    dbc.TableName := ExtractFileName(EditCaminhoDBF.Text);
    dbc.Open;


    ComboNome.Clear;
    ComboDataCx.Clear;
    ComboVenc.Clear;
    ComboAtraso.Clear;
    ComboValor.Clear;
    ComboJuros.Clear;
    ComboValorConta.Clear;


    for i := 0 to dbc.FieldDefs.Count - 1 do
    begin
      ComboNome.Items.Add(dbc.FieldDefs[i].Name);
      ComboDataCx.Items.Add(dbc.FieldDefs[i].Name);
      ComboVenc.Items.Add(dbc.FieldDefs[i].Name);
      ComboAtraso.Items.Add(dbc.FieldDefs[i].Name);
      ComboValor.Items.Add(dbc.FieldDefs[i].Name);
      ComboJuros.Items.Add(dbc.FieldDefs[i].Name);
      ComboValorConta.Items.Add(dbc.FieldDefs[i].Name);
    end;


    ComboNome.Text := 'NOME';
    ComboDataCx.Text := 'DATACX';
    ComboVenc.Text := 'VENC';
    ComboAtraso.Text := 'ATRASO';
    ComboValor.Text := 'VALOR';
    ComboJuros.Text := 'N_JUROS';
    ComboValorConta.Text := 'VALORCONTA';

    ShowMessage('Colunas de Contas a Receber carregadas com sucesso!');
  finally
    dbc.Close;
    dbc.Free;
  end;
end;

{ MÉTODOS ESPECÍFICOS DE CONTAS }

function TForm4.ParseDateBR(const S: string): TDateTime;
var
  FS: TFormatSettings;
  T: string;
begin
  FS := DefaultFormatSettings;
  FS.DateSeparator := '/';
  FS.ShortDateFormat := 'dd/mm/yyyy';
  T := Trim(S);
  if not TryStrToDate(T, Result, FS) then
    Result := Now;
end;


function TForm4.ParseFloatBR(const S: string): Double;
var
  T: string;
  FS: TFormatSettings;
  i: Integer;
  hasComma, hasDot: Boolean;
begin
  // Corrige valores vindos do DBF
  T := Trim(S);

  // Remove aspas se existirem
  T := StringReplace(T, '"', '', [rfReplaceAll]);

  // Caso venha vazio, retorna 0
  if T = '' then
  begin
    Result := 0;
    Exit;
  end;

  // Remove símbolos monetários
  T := StringReplace(T, 'R$', '', [rfReplaceAll, rfIgnoreCase]);
  T := StringReplace(T, ' ', '', [rfReplaceAll]);
  T := StringReplace(T, #9, '', [rfReplaceAll]);

  // Verifica se tem vírgula ou ponto
  hasComma := Pos(',', T) > 0;
  hasDot := Pos('.', T) > 0;

  // Configura formato brasileiro como padrão
  FS := DefaultFormatSettings;
  FS.DecimalSeparator := ',';
  FS.ThousandSeparator := '.';

  if hasComma and hasDot then
  begin
    // Tem ambos: assume que vírgula é decimal e ponto é milhar
    T := StringReplace(T, '.', '', [rfReplaceAll]);
  end
  else if hasDot and (not hasComma) then
  begin
    // Só tem ponto: assume que é decimal (formato internacional)
    FS.DecimalSeparator := '.';
    FS.ThousandSeparator := #0;
  end;
  // Se só tem vírgula, mantém formato brasileiro padrão

  // Tenta converter
  if not TryStrToFloat(T, Result, FS) then
  begin
    // Se falhar, tenta formato padrão do sistema como fallback
    if not TryStrToFloat(T, Result) then
    begin
      Result := 0;
    end;
  end;
end;

procedure TForm4.WarmUpFaturaCounter(Q: TSQLQuery);
begin
  Prep(Q, 'SELECT COALESCE(MAX(IDFATURA), 0) AS N FROM FATURAS');
  Q.Open;
  LastIDFatura := Q.FieldByName('N').AsInteger;
  Q.Close;
end;

function TForm4.NextIDFatura(Q: TSQLQuery): Integer;
begin
  Inc(LastIDFatura);
  Result := LastIDFatura;
end;

{ MÉTODO PRINCIPAL DE IMPORTAÇÃO }

procedure TForm4.ImportarContasReceber;
var
  dbfC: TDbf;
  total, atual, puladas, ok: Integer;
  sNome, sDataCx, sVenc, sAtraso, sValor, sJuros, sValorConta: string;
  dtIncl, dtVenc: TDateTime;
  vValor, vJuros, vValorConta: Double;
  codLoja, codTerminal: Integer;
  codCarteira, codPlano: String;
  idNatDoc: Variant;
  idFatura, codCliente: Integer;
  nomeLookup: string;
  connLocal: TIBConnection;
  transLocal: TSQLTransaction;
  queryLocal: TSQLQuery;
  sLog: TStringList;
begin

  codLoja := StrToIntDef(Trim(EditCodLoja.Text), 1);
  codTerminal := StrToIntDef(Trim(EditCodTerminal.Text), 1);

  if Trim(EditIdNatureza.Text) = '' then
    idNatDoc := Null
  else
    idNatDoc := StrToIntDef(Trim(EditIdNatureza.Text), 1);

  codPlano := EditCodPlanoPagto.Text;
  codCarteira := EditCodCarteira.Text;


  connLocal := TIBConnection.Create(nil);
  transLocal := TSQLTransaction.Create(nil);
  queryLocal := TSQLQuery.Create(nil);
  dbfC := TDbf.Create(nil);
  sLog := TStringList.Create;

  try

    LabelProgresso.Caption := 'Abrindo arquivo DBF...';
    Application.ProcessMessages;

    dbfC.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
    dbfC.TableName := ExtractFileName(EditCaminhoDBF.Text);
    dbfC.Open;


    LabelProgresso.Caption := 'Conectando ao banco...';
    Application.ProcessMessages;

    connLocal.DatabaseName := EditCaminhoFDB.Text;
    connLocal.UserName := 'sysdba';
    connLocal.Password := 'masterkey';
    connLocal.CharSet := 'UTF8';
    connLocal.Params.Add('lc_ctype=UTF8');
    connLocal.Transaction := transLocal;
    transLocal.DataBase := connLocal;
    queryLocal.DataBase := connLocal;
    queryLocal.Transaction := transLocal;

    connLocal.Open;
    transLocal.StartTransaction;


    LabelProgresso.Caption := 'Buscando último ID...';
    Application.ProcessMessages;
    WarmUpFaturaCounter(queryLocal);

    total := dbfC.RecordCount;
    atual := 0;
    puladas := 0;
    ok := 0;
    ProgressBar.Position := 0;
    LabelProgresso.Caption := 'Iniciando importação...';
    Application.ProcessMessages;

    sLog.Add('=== LOG DE IMPORTAÇÃO - ' + FormatDateTime('dd/mm/yyyy hh:nn:ss', Now) + ' ===');
    sLog.Add('');

    dbfC.First;
    while not dbfC.EOF do
    begin
      Inc(atual);

      try

        sNome := PrepareStringForDB(Trim(dbfC.FieldByName(ComboNome.Text).AsString));
        sDataCx := Trim(dbfC.FieldByName(ComboDataCx.Text).AsString);
        sVenc := Trim(dbfC.FieldByName(ComboVenc.Text).AsString);
        sAtraso := Trim(dbfC.FieldByName(ComboAtraso.Text).AsString);
        sValor := Trim(dbfC.FieldByName(ComboValor.Text).AsString);
        sJuros := Trim(dbfC.FieldByName(ComboJuros.Text).AsString);
        sValorConta := Trim(dbfC.FieldByName(ComboValorConta.Text).AsString);


        if (sNome = '') or (sVenc = '') or (sValor = '') then
        begin
          sLog.Add(Format('❌ Pulado %d: Campo essencial vazio', [atual]));
          Inc(puladas);
          dbfC.Next;
          Continue;
        end;


        try
          dtIncl := ParseDateBR(sDataCx);
          dtVenc := ParseDateBR(sVenc);
        except
          sLog.Add(Format('❌ Pulado %d: Data inválida', [atual]));
          Inc(puladas);
          dbfC.Next;
          Continue;
        end;


        try
  vValor := ParseFloatBR(sValor);
  vJuros := ParseFloatBR(sJuros);
  vValorConta := ParseFloatBR(sValorConta);

  // DEBUG - Mostra os valores que estão sendo processados
  sLog.Add(Format('DEBUG %d: Valor=%s -> %f, Juros=%s -> %f, ValorConta=%s -> %f',
    [atual, sValor, vValor, sJuros, vJuros, sValorConta, vValorConta]));

  // LÓGICA CORRIGIDA: Prioriza VALORCONTA, depois VALOR, depois JUROS
  if vValorConta > 0 then
  begin
    // Se VALORCONTA é válido, usa ele
    if vValor = 0 then
      vValor := vValorConta - vJuros;
  end
  else if vValor > 0 then
  begin
    // Se VALOR é válido, calcula VALORCONTA
    vValorConta := vValor + vJuros;
  end
  else if vJuros > 0 then
  begin
    // Se ambos VALOR e VALORCONTA são zero mas JUROS é positivo, usa JUROS como valor
    vValor := vJuros;
    vValorConta := vJuros;
    vJuros := 0; // Zera os juros já que foi incorporado ao valor principal
    sLog.Add(Format('💡 Ajustado %d: Usando juros como valor principal', [atual]));
  end
  else
  begin
    // Se todos são zero, pula
    sLog.Add(Format('❌ Pulado %d: Valor, ValorConta e Juros são zero', [atual]));
    Inc(puladas);
    dbfC.Next;
    Continue;
  end;

  // Validação final - aceita qualquer valor positivo
  if (vValorConta <= 0) then
  begin
    sLog.Add(Format('❌ Pulado %d: ValorConta inválido: %f', [atual, vValorConta]));
    Inc(puladas);
    dbfC.Next;
    Continue;
  end;

except
  on E: Exception do
  begin
    sLog.Add(Format('❌ Erro converter valores %d: %s', [atual, E.Message]));
    Inc(puladas);
    dbfC.Next;
    Continue;
  end;
end;

        codCliente := 0;
        nomeLookup := UpperCase(Trim(sNome));

        queryLocal.Close;
        queryLocal.SQL.Text := 'SELECT CODCLIENTE FROM CLIENTE WHERE UPPER(TRIM(NOME)) = ' + QuotedStr(nomeLookup);
        queryLocal.Open;

        if not queryLocal.IsEmpty then
          codCliente := queryLocal.FieldByName('CODCLIENTE').AsInteger;

        queryLocal.Close;

        if codCliente = 0 then
        begin
          sLog.Add(Format('❌ Pulado %d: Cliente não encontrado', [atual]));
          Inc(puladas);
          dbfC.Next;
          Continue;
        end;


        idFatura := NextIDFatura(queryLocal);



        queryLocal.Close;
        queryLocal.SQL.Text :=
          'INSERT INTO FATURAS (' +
          'IDFATURA, R_P, DATA_INCLUSAO, VALOR, OBS, ORIGEM_FATURA, ' +
          'IDFATURAMENTO, IDFATURARENEGOCIADA, CODTERMINAL, CODLOJA, ' +
          'ALTERADO, DESCONTO, VALOR_ORIG, ESTADO, CODLOJARENEGOCIADA, CODTERMINALRENEGOCIADA) ' +
          'VALUES (' +
          ':IDFATURA, :R_P, :DATA_INCLUSAO, :VALOR, :OBS, :ORIGEM_FATURA, ' +
          ':IDFATURAMENTO, :IDFATURARENEGOCIADA, :CODTERMINAL, :CODLOJA, ' +
          ':ALTERADO, :DESCONTO, :VALOR_ORIG, :ESTADO, :CODLOJARENEGOCIADA, :CODTERMINALRENEGOCIADA)';

        queryLocal.ParamByName('IDFATURA').AsInteger := idFatura;
        queryLocal.ParamByName('R_P').AsString := 'R';
        queryLocal.ParamByName('DATA_INCLUSAO').AsDate := dtIncl;
        queryLocal.ParamByName('VALOR').AsFloat := vValorConta;
        queryLocal.ParamByName('OBS').AsString := '';
        queryLocal.ParamByName('ORIGEM_FATURA').AsString := 'I';


        queryLocal.ParamByName('IDFATURAMENTO').Clear;
        queryLocal.ParamByName('IDFATURARENEGOCIADA').Clear;

        queryLocal.ParamByName('CODTERMINAL').AsInteger := codTerminal;
        queryLocal.ParamByName('CODLOJA').AsInteger := codLoja;
        queryLocal.ParamByName('ALTERADO').AsDateTime := Now;
        queryLocal.ParamByName('DESCONTO').AsFloat := 0;
        queryLocal.ParamByName('VALOR_ORIG').AsFloat := vValor;
        queryLocal.ParamByName('ESTADO').AsString := 'A';
        queryLocal.ParamByName('CODLOJARENEGOCIADA').Clear;
        queryLocal.ParamByName('CODTERMINALRENEGOCIADA').Clear;

        queryLocal.ExecSQL;


        queryLocal.Close;
        queryLocal.SQL.Text :=
          'INSERT INTO FATURAS_RECEBER (' +
          'IDFATURA, CODCLIENTE, ID_NATUREZA_DOCUMENTO, CODPLANOPAGTO, CODTERMINAL, CODLOJA, ALTERADO) ' +
          'VALUES (' +
          ':IDFATURA, :CODCLIENTE, :ID_NATUREZA_DOCUMENTO, :CODPLANOPAGTO, :CODTERMINAL, :CODLOJA, :ALTERADO)';

        queryLocal.ParamByName('IDFATURA').AsInteger := idFatura;
        queryLocal.ParamByName('CODCLIENTE').AsInteger := codCliente;


        if VarIsNull(idNatDoc) then
          queryLocal.ParamByName('ID_NATUREZA_DOCUMENTO').Clear
        else
          queryLocal.ParamByName('ID_NATUREZA_DOCUMENTO').AsInteger := idNatDoc;

        queryLocal.ParamByName('CODPLANOPAGTO').AsString := codPlano;
        queryLocal.ParamByName('CODTERMINAL').AsInteger := codTerminal;
        queryLocal.ParamByName('CODLOJA').AsInteger := codLoja;
        queryLocal.ParamByName('ALTERADO').AsDateTime := Now;

        queryLocal.ExecSQL;


        queryLocal.Close;
        queryLocal.SQL.Text :=
          'INSERT INTO FATURAS_PARCELAS (' +
          'IDFATURA, PARCELA, VENCIMENTO, VALOR, DATA_LIQUIDADO, CODCARTEIRA, ' +
          'DIAS_ANTECIPACAO, PORC_DESC_ANTECIPACAO, DIA_CARENCIA, PORC_MULTA, ' +
          'PORC_MORA_DIA, PORC_PAGTO_MINIMO, NUMERO_BANCARIO, TIPO_OPERACAO, ' +
          'CODREMESSA, QTD_REMESSA_GERADA, QTD_BOLETO_IMPRESSO, QTD_ENVIO_EMAIL, ' +
          'CODTERMINAL, CODLOJA, ALTERADO, TIPO_JUROS, ULT_NOTIFICACAO) ' +
          'VALUES (' +
          ':IDFATURA, :PARCELA, :VENCIMENTO, :VALOR, :DATA_LIQUIDADO, :CODCARTEIRA, ' +
          ':DIAS_ANTECIPACAO, :PORC_DESC_ANTECIPACAO, :DIA_CARENCIA, :PORC_MULTA, ' +
          ':PORC_MORA_DIA, :PORC_PAGTO_MINIMO, :NUMERO_BANCARIO, :TIPO_OPERACAO, ' +
          ':CODREMESSA, :QTD_REMESSA_GERADA, :QTD_BOLETO_IMPRESSO, :QTD_ENVIO_EMAIL, ' +
          ':CODTERMINAL, :CODLOJA, :ALTERADO, :TIPO_JUROS, :ULT_NOTIFICACAO)';

        queryLocal.ParamByName('IDFATURA').AsInteger := idFatura;
        queryLocal.ParamByName('PARCELA').AsInteger := 1;
        queryLocal.ParamByName('VENCIMENTO').AsDate := dtVenc;
        queryLocal.ParamByName('VALOR').AsFloat := vValorConta;
        queryLocal.ParamByName('DATA_LIQUIDADO').Clear;
        queryLocal.ParamByName('CODCARTEIRA').AsString := codCarteira;
        queryLocal.ParamByName('DIAS_ANTECIPACAO').AsInteger := 0;
        queryLocal.ParamByName('PORC_DESC_ANTECIPACAO').AsFloat := 0;
        queryLocal.ParamByName('DIA_CARENCIA').AsInteger := 0;
        queryLocal.ParamByName('PORC_MULTA').AsFloat := 0;
        queryLocal.ParamByName('PORC_MORA_DIA').AsFloat := 0;
        queryLocal.ParamByName('PORC_PAGTO_MINIMO').AsFloat := 0;
        queryLocal.ParamByName('NUMERO_BANCARIO').Clear;
        queryLocal.ParamByName('TIPO_OPERACAO').Clear;
        queryLocal.ParamByName('CODREMESSA').Clear;
        queryLocal.ParamByName('QTD_REMESSA_GERADA').AsInteger := 0;
        queryLocal.ParamByName('QTD_BOLETO_IMPRESSO').AsInteger := 0;
        queryLocal.ParamByName('QTD_ENVIO_EMAIL').AsInteger := 0;
        queryLocal.ParamByName('CODTERMINAL').AsInteger := codTerminal;
        queryLocal.ParamByName('CODLOJA').AsInteger := codLoja;
        queryLocal.ParamByName('ALTERADO').AsDateTime := Now;
        queryLocal.ParamByName('TIPO_JUROS').Clear;
        queryLocal.ParamByName('ULT_NOTIFICACAO').Clear;

        queryLocal.ExecSQL;

        Inc(ok);
        sLog.Add(Format('✅ Inserido %d: %s - R$ %.2f', [atual, sNome, vValorConta]));


        if (atual mod 100) = 0 then
        begin
          transLocal.CommitRetaining;
          LabelProgresso.Caption := Format('Importando %d/%d... (ok=%d | pulados=%d)',
            [atual, total, ok, puladas]);
          Application.ProcessMessages;
        end;

        if total > 0 then
          ProgressBar.Position := Round((atual / total) * 100);

      except
        on E: Exception do
        begin
          sLog.Add(Format('❌ Erro geral %d: %s', [atual, E.Message]));
          Inc(puladas);


          if transLocal.Active then
            transLocal.RollbackRetaining;
        end;
      end;

      dbfC.Next;
    end;


    if transLocal.Active then
      transLocal.Commit;


    sLog.Add('');
    sLog.Add('=== RESUMO DA IMPORTAÇÃO ===');
    sLog.Add(Format('Total processado: %d', [total]));
    sLog.Add(Format('Inseridos com sucesso: %d', [ok]));
    sLog.Add(Format('Registros pulados: %d', [puladas]));

    if total > 0 then
      sLog.Add(Format('Taxa de sucesso: %.2f%%', [(ok / total) * 100]))
    else
      sLog.Add('Taxa de sucesso: 0%');


    sLog.SaveToFile('log_importacao_contas_' + FormatDateTime('ddmmyyyy_hhnnss', Now) + '.txt');


    LabelProgresso.Caption := Format('✅ Concluído: %d inseridos | %d pulados', [ok, puladas]);
    ProgressBar.Position := 100;
    ShowMessage(LabelProgresso.Caption);

  finally

    sLog.Free;
    FreeAndNil(queryLocal);
    FreeAndNil(transLocal);
    FreeAndNil(connLocal);
    if Assigned(dbfC) then
    begin
      if dbfC.Active then dbfC.Close;
      FreeAndNil(dbfC);
    end;
  end;
end;


end.
