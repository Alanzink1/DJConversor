unit uImportarClientes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, Variants;

type
  { TForm3 }
  TForm3 = class(TForm)
    ButtonBuscarDBF: TButton;
    ButtonBuscarFDB: TButton;
    ButtonImportar: TButton;
    EditCaminhoDBF: TEdit;
    EditCaminhoFDB: TEdit;
    LabelDBF: TLabel;
    LabelFDB: TLabel;
    LabelNome: TLabel;
    LabelTipo: TLabel;
    LabelCpfCnpj: TLabel;
    LabelCidade: TLabel;
    LabelProgresso: TLabel;
    ComboNome: TComboBox;
    ComboTipo: TComboBox;
    ComboCpfCnpj: TComboBox;
    ComboCidade: TComboBox;
    ProgressBar: TProgressBar;
    OpenDialog: TOpenDialog;
    LabelContribuinte: TLabel;
    ComboContribuinte: TComboBox;
    LabelOpcaoContribuinte: TLabel;
    ComboOpcaoContribuinte: TComboBox;

    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure ComboOpcaoContribuinteChange(Sender: TObject);

  private
    FPortaFirebird: Integer;
    procedure CarregarColunas;
    procedure ImportarClientes;
    procedure LogMsg(const nomeArq, msg: string);
    function RemoveAcentos(const S: string): string;
    function PrepareStringForDB(const S: string): string;
    function DeterminarContribuinteICMS(const cpfCnpj: string; const colunaContribuinte: string; dbfC: TDbf): Integer;

  public
    property PortaFirebird: Integer read FPortaFirebird write FPortaFirebird;
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  Form3: TForm3;

implementation

{$R *.lfm}

const
  LOGS_ENABLED = True;

{ TForm3 }

constructor TForm3.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TForm3.Destroy;
begin
  inherited Destroy;
end;

procedure TForm3.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor - Clientes';
  Width := 850;
  Height := 600;
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
  LabelNome.Caption := 'Nome:';
  LabelNome.Left := 20;
  LabelNome.Top := 110;

  ComboNome := TComboBox.Create(Self);
  ComboNome.Parent := Self;
  ComboNome.Left := 150;
  ComboNome.Top := 106;
  ComboNome.Width := 200;

  LabelTipo := TLabel.Create(Self);
  LabelTipo.Parent := Self;
  LabelTipo.Caption := 'Tipo Pessoa:';
  LabelTipo.Left := 20;
  LabelTipo.Top := 150;

  ComboTipo := TComboBox.Create(Self);
  ComboTipo.Parent := Self;
  ComboTipo.Left := 150;
  ComboTipo.Top := 146;
  ComboTipo.Width := 200;

  LabelCpfCnpj := TLabel.Create(Self);
  LabelCpfCnpj.Parent := Self;
  LabelCpfCnpj.Caption := 'CPF/CNPJ:';
  LabelCpfCnpj.Left := 20;
  LabelCpfCnpj.Top := 190;

  ComboCpfCnpj := TComboBox.Create(Self);
  ComboCpfCnpj.Parent := Self;
  ComboCpfCnpj.Left := 150;
  ComboCpfCnpj.Top := 186;
  ComboCpfCnpj.Width := 200;


  LabelCidade := TLabel.Create(Self);
  LabelCidade.Parent := Self;
  LabelCidade.Caption := 'Cidade:';
  LabelCidade.Left := 400;
  LabelCidade.Top := 110;

  ComboCidade := TComboBox.Create(Self);
  ComboCidade.Parent := Self;
  ComboCidade.Left := 500;
  ComboCidade.Top := 106;
  ComboCidade.Width := 200;

  LabelContribuinte := TLabel.Create(Self);
  LabelContribuinte.Parent := Self;
  LabelContribuinte.Caption := 'Coluna Contribuinte:';
  LabelContribuinte.Left := 400;
  LabelContribuinte.Top := 150;

  ComboContribuinte := TComboBox.Create(Self);
  ComboContribuinte.Parent := Self;
  ComboContribuinte.Left := 500;
  ComboContribuinte.Top := 146;
  ComboContribuinte.Width := 200;


  LabelOpcaoContribuinte := TLabel.Create(Self);
  LabelOpcaoContribuinte.Parent := Self;
  LabelOpcaoContribuinte.Caption := 'Configuração do Contribuinte ICMS:';
  LabelOpcaoContribuinte.Left := 20;
  LabelOpcaoContribuinte.Top := 230;

  ComboOpcaoContribuinte := TComboBox.Create(Self);
  ComboOpcaoContribuinte.Parent := Self;
  ComboOpcaoContribuinte.Left := 250;
  ComboOpcaoContribuinte.Top := 226;
  ComboOpcaoContribuinte.Width := 450;
  ComboOpcaoContribuinte.Items.Add('Todos como NÃO CONTRIBUINTES (9)');
  ComboOpcaoContribuinte.Items.Add('Todos como CONTRIBUINTES (1)');
  ComboOpcaoContribuinte.Items.Add('Todos como ISENTOS (2)');
  ComboOpcaoContribuinte.Items.Add('Automático: CNPJ=Contribuinte(1), CPF=Não Contribuinte(9)');
  ComboOpcaoContribuinte.Items.Add('Usar coluna específica do DBF');
  ComboOpcaoContribuinte.ItemIndex := 0;
  ComboOpcaoContribuinte.OnChange := @ComboOpcaoContribuinteChange;

  ButtonImportar := TButton.Create(Self);
  ButtonImportar.Parent := Self;
  ButtonImportar.Caption := '🚀 Importar Clientes';
  ButtonImportar.Left := 20;
  ButtonImportar.Top := 280;
  ButtonImportar.Width := 200;
  ButtonImportar.Height := 40;
  ButtonImportar.OnClick := @ButtonImportarClick;

  ProgressBar := TProgressBar.Create(Self);
  ProgressBar.Parent := Self;
  ProgressBar.Left := 20;
  ProgressBar.Top := 340;
  ProgressBar.Width := 680;
  ProgressBar.Min := 0;
  ProgressBar.Max := 100;

  LabelProgresso := TLabel.Create(Self);
  LabelProgresso.Parent := Self;
  LabelProgresso.Left := 20;
  LabelProgresso.Top := 370;
  LabelProgresso.Caption := '';

  OpenDialog := TOpenDialog.Create(Self);
end;

procedure TForm3.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
end;

procedure TForm3.ComboOpcaoContribuinteChange(Sender: TObject);
begin

  ComboContribuinte.Enabled := (ComboOpcaoContribuinte.ItemIndex = 4);
  LabelContribuinte.Enabled := (ComboOpcaoContribuinte.ItemIndex = 4);
end;

{ MÉTODOS UTILITÁRIOS }

procedure TForm3.LogMsg(const nomeArq, msg: string);
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

function TForm3.RemoveAcentos(const S: string): string;
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

function TForm3.PrepareStringForDB(const S: string): string;
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

function TForm3.DeterminarContribuinteICMS(const cpfCnpj: string; const colunaContribuinte: string; dbfC: TDbf): Integer;
var
  cpfCnpjLimpo: string;
  valorColuna: string;
begin
  cpfCnpjLimpo := StringReplace(cpfCnpj, '.', '', [rfReplaceAll]);
  cpfCnpjLimpo := StringReplace(cpfCnpjLimpo, '-', '', [rfReplaceAll]);
  cpfCnpjLimpo := StringReplace(cpfCnpjLimpo, '/', '', [rfReplaceAll]);

  case ComboOpcaoContribuinte.ItemIndex of
    0: Result := 9;
    1: Result := 1;
    2: Result := 2;

    3: begin
      if Length(cpfCnpjLimpo) = 14 then
        Result := 1
      else if Length(cpfCnpjLimpo) = 11 then
        Result := 9
      else
        Result := 9;
    end;

    4: begin
      if colunaContribuinte <> '' then
      begin
        valorColuna := UpperCase(Trim(dbfC.FieldByName(colunaContribuinte).AsString));

        if (valorColuna = '1') or (valorColuna = 'S') or (valorColuna = 'SIM') or
           (valorColuna = 'CONTRIBUINTE') or (valorColuna = 'C') then
          Result := 1
        else if (valorColuna = '2') or (valorColuna = 'I') or (valorColuna = 'ISENTO') then
          Result := 2
        else if (valorColuna = '9') or (valorColuna = 'N') or (valorColuna = 'NAO') or
                (valorColuna = 'NÃO') or (valorColuna = 'NÃO CONTRIBUINTE') then
          Result := 9
        else
          Result := 9;
      end
      else
        Result := 9;
    end;

  else
    Result := 9;
  end;
end;

{ MÉTODOS DE UI }

procedure TForm3.ButtonBuscarDBFClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Arquivos DBF|*.dbf';
  if OpenDialog.Execute then
  begin
    EditCaminhoDBF.Text := OpenDialog.FileName;
    CarregarColunas;
  end;
end;

procedure TForm3.ButtonBuscarFDBClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Banco Firebird|*.fdb';
  if OpenDialog.Execute then
    EditCaminhoFDB.Text := OpenDialog.FileName;
end;

procedure TForm3.ButtonImportarClick(Sender: TObject);
begin
  if (EditCaminhoDBF.Text = '') or (EditCaminhoFDB.Text = '') then
  begin
    ShowMessage('Selecione os arquivos DBF e FDB antes de importar clientes.');
    Exit;
  end;

  if (ComboNome.Text = '') or (ComboCpfCnpj.Text = '') or (ComboCidade.Text = '') then
  begin
    ShowMessage('Mapeie as colunas obrigatórias: Nome, CPF/CNPJ e Cidade.');
    Exit;
  end;

  if (ComboOpcaoContribuinte.ItemIndex = 4) and (ComboContribuinte.Text = '') then
  begin
    ShowMessage('Quando selecionar "Usar coluna específica", mapeie a coluna correspondente.');
    Exit;
  end;

  ImportarClientes;
end;

procedure TForm3.CarregarColunas;
var
  i: Integer;
  dbc: TDbf;
begin
  if not FileExists(EditCaminhoDBF.Text) then
    Exit;

  dbc := TDbf.Create(nil);
  try
    dbc.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
    dbc.TableName := ExtractFileName(EditCaminhoDBF.Text);
    dbc.Open;

    ComboNome.Clear;
    ComboTipo.Clear;
    ComboCpfCnpj.Clear;
    ComboCidade.Clear;
    ComboContribuinte.Clear;

    for i := 0 to dbc.FieldDefs.Count - 1 do
    begin
      ComboNome.Items.Add(dbc.FieldDefs[i].Name);
      ComboTipo.Items.Add(dbc.FieldDefs[i].Name);
      ComboCpfCnpj.Items.Add(dbc.FieldDefs[i].Name);
      ComboCidade.Items.Add(dbc.FieldDefs[i].Name);
      ComboContribuinte.Items.Add(dbc.FieldDefs[i].Name);
    end;

    ShowMessage('Colunas de clientes carregadas com sucesso!');
  finally
    dbc.Close;
    dbc.Free;
  end;
end;

{ MÉTODO PRINCIPAL DE IMPORTAÇÃO DE CLIENTES }


{ MÉTODO PRINCIPAL DE IMPORTAÇÃO DE CLIENTES }


{ MÉTODO PRINCIPAL DE IMPORTAÇÃO DE CLIENTES }


procedure TForm3.ImportarClientes;
var
  dbfC: TDbf;
  connC: TIBConnection;
  transC: TSQLTransaction;
  queryC: TSQLQuery;
  codigo, puladas, total, atual: Integer;
  nome, tipo_pessoa, cidade, cpf_cnpj, cpais: string;
  cadastrado_em: TDateTime;
  contribuinteICMS: Integer;
  colunaContribuinte: string;
begin
  cpais := '1058';
  codigo := 0;
  puladas := 0;


  if ComboOpcaoContribuinte.ItemIndex = 4 then
    colunaContribuinte := ComboContribuinte.Text
  else
    colunaContribuinte := '';

  dbfC := TDbf.Create(nil);
  connC := TIBConnection.Create(nil);
  transC := TSQLTransaction.Create(nil);
  queryC := TSQLQuery.Create(nil);

  try

    LabelProgresso.Caption := 'Abrindo arquivo DBF...';
    Application.ProcessMessages;

    dbfC.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
    dbfC.TableName := ExtractFileName(EditCaminhoDBF.Text);
    dbfC.Open;


    LabelProgresso.Caption := 'Conectando ao banco...';
    Application.ProcessMessages;

    connC.DatabaseName := EditCaminhoFDB.Text;
    connC.UserName := 'sysdba';
    connC.Password := 'masterkey';
    connC.CharSet := 'UTF8';
    connC.Params.Add('lc_ctype=UTF8');
    connC.Transaction := transC;
    transC.DataBase := connC;
    queryC.DataBase := connC;
    queryC.Transaction := transC;

    connC.Open;


    queryC.SQL.Text := 'SELECT COALESCE(MAX(CODCLIENTE), 0) AS ULTIMO FROM CLIENTE';
    queryC.Open;
    codigo := queryC.FieldByName('ULTIMO').AsInteger;
    queryC.Close;


    LogMsg('log_diagnostico_detalhado.txt', '=== INICIANDO DIAGNÓSTICO DETALHADO ===');


    try
      queryC.SQL.Text := 'INSERT INTO CLIENTE (CODCLIENTE, NOME, F_J, CADASTRADO_EM, NIVEL_CREDITO) VALUES (999999, ''TESTE'', ''F'', CURRENT_TIMESTAMP, 5)';
      queryC.ExecSQL;
      queryC.SQL.Text := 'DELETE FROM CLIENTE WHERE CODCLIENTE = 999999';
      queryC.ExecSQL;
      LogMsg('log_diagnostico_detalhado.txt', '✅ Teste mínimo (COD, NOME, F_J, CADASTRADO_EM) - OK');
    except
      on E: Exception do
        LogMsg('log_diagnostico_detalhado.txt', '❌ Teste mínimo falhou: ' + E.Message);
    end;


    try
      queryC.SQL.Text := 'INSERT INTO CLIENTE (CODCLIENTE, NOME, F_J, CPF_CNPJ, CADASTRADO_EM, NIVEL_CREDITO) VALUES (999998, ''TESTE'', ''F'', ''12345678901'', CURRENT_TIMESTAMP, 5)';
      queryC.ExecSQL;
      queryC.SQL.Text := 'DELETE FROM CLIENTE WHERE CODCLIENTE = 999998';
      queryC.ExecSQL;
      LogMsg('log_diagnostico_detalhado.txt', '✅ Teste com CPF_CNPJ - OK');
    except
      on E: Exception do
        LogMsg('log_diagnostico_detalhado.txt', '❌ Teste com CPF_CNPJ falhou: ' + E.Message);
    end;


    try
      queryC.SQL.Text := 'INSERT INTO CLIENTE (CODCLIENTE, NOME, F_J, CPF_CNPJ, CIDADE, CADASTRADO_EM, NIVEL_CREDITO) VALUES (999997, ''TESTE'', ''F'', ''12345678901'', ''TESTE'', CURRENT_TIMESTAMP, 5)';
      queryC.ExecSQL;
      queryC.SQL.Text := 'DELETE FROM CLIENTE WHERE CODCLIENTE = 999997';
      queryC.ExecSQL;
      LogMsg('log_diagnostico_detalhado.txt', '✅ Teste com CIDADE - OK');
    except
      on E: Exception do
        LogMsg('log_diagnostico_detalhado.txt', '❌ Teste com CIDADE falhou: ' + E.Message);
    end;


    try
      queryC.SQL.Text := 'INSERT INTO CLIENTE (CODCLIENTE, NOME, F_J, CPF_CNPJ, CIDADE, CONTRIBUINTEICMS, CADASTRADO_EM, NIVEL_CREDITO) VALUES (999996, ''TESTE'', ''F'', ''12345678901'', ''TESTE'', 9, CURRENT_TIMESTAMP, 5)';
      queryC.ExecSQL;
      queryC.SQL.Text := 'DELETE FROM CLIENTE WHERE CODCLIENTE = 999996';
      queryC.ExecSQL;
      LogMsg('log_diagnostico_detalhado.txt', '✅ Teste com CONTRIBUINTEICMS - OK');
    except
      on E: Exception do
        LogMsg('log_diagnostico_detalhado.txt', '❌ Teste com CONTRIBUINTEICMS falhou: ' + E.Message);
    end;


    try
      queryC.SQL.Text := 'INSERT INTO CLIENTE (' +
        'CODCLIENTE, NOME, CADASTRADO_EM, CIDADE, F_J, CPF_CNPJ, ' +
        'NIVEL_CREDITO, CPAIS, CONTRIBUINTEICMS, ALTERADO, DEBITO' +
        ') VALUES (' +
        '999995, ''TESTE COMPLETO'', CURRENT_TIMESTAMP, ''TESTE'', ''F'', ''12345678901'', ' +
        '5, ''1058'', 9, CURRENT_TIMESTAMP, 0)';
      queryC.ExecSQL;
      queryC.SQL.Text := 'DELETE FROM CLIENTE WHERE CODCLIENTE = 999995';
      queryC.ExecSQL;
      LogMsg('log_diagnostico_detalhado.txt', '✅ Teste completo - OK');
    except
      on E: Exception do
        LogMsg('log_diagnostico_detalhado.txt', '❌ Teste completo falhou: ' + E.Message);
    end;


    if transC.Active then
      transC.Commit;
    transC.StartTransaction;


    total := dbfC.RecordCount;
    atual := 0;
    ProgressBar.Position := 0;

    LabelProgresso.Caption := Format('Iniciando importação de %d clientes...', [total]);
    Application.ProcessMessages;

    dbfC.First;
    while not dbfC.EOF do
    begin
      Inc(codigo);
      Inc(atual);


      try
        nome := PrepareStringForDB(UTF8Encode(Trim(dbfC.FieldByName(ComboNome.Text).AsString)));

        if ComboTipo.Text <> '' then
          tipo_pessoa := Trim(dbfC.FieldByName(ComboTipo.Text).AsString)
        else
          tipo_pessoa := '';

        cpf_cnpj := PrepareStringForDB(Trim(dbfC.FieldByName(ComboCpfCnpj.Text).AsString));
        cidade := PrepareStringForDB(UTF8Encode(Trim(dbfC.FieldByName(ComboCidade.Text).AsString)));
        cadastrado_em := Now;
      except
        on E: Exception do
        begin
          LogMsg('log_erros_clientes.txt', 'Erro ao ler dados DBF: ' + E.Message);
          Inc(puladas);
          Dec(codigo);
          dbfC.Next;
          Continue;
        end;
      end;


      if (nome = '') or (cpf_cnpj = '') or (cidade = '') then
      begin
        LogMsg('log_erros_clientes.txt',
          Format('Cliente pulado - Campos obrigatórios vazios: Nome="%s" CPF/CNPJ="%s" Cidade="%s"',
          [nome, cpf_cnpj, cidade]));
        Inc(puladas);
        Dec(codigo);
        dbfC.Next;
        Continue;
      end;


      if tipo_pessoa <> '' then
      begin
        tipo_pessoa := UpperCase(Trim(tipo_pessoa));

        if Length(tipo_pessoa) > 1 then
          tipo_pessoa := tipo_pessoa[1];


        if (tipo_pessoa = 'F') or (tipo_pessoa = '1') then
          tipo_pessoa := 'F'
        else if (tipo_pessoa = 'J') or (tipo_pessoa = '2') then
          tipo_pessoa := 'J'
        else
        begin

          tipo_pessoa := '';
        end;
      end;


      if (tipo_pessoa = '') or ((tipo_pessoa <> 'F') and (tipo_pessoa <> 'J')) then
      begin

        cpf_cnpj := StringReplace(cpf_cnpj, '.', '', [rfReplaceAll]);
        cpf_cnpj := StringReplace(cpf_cnpj, '-', '', [rfReplaceAll]);
        cpf_cnpj := StringReplace(cpf_cnpj, '/', '', [rfReplaceAll]);

        if Length(cpf_cnpj) = 14 then
          tipo_pessoa := 'J'
        else if Length(cpf_cnpj) = 11 then
          tipo_pessoa := 'F'
        else
          tipo_pessoa := 'F';
      end;


      try
        contribuinteICMS := DeterminarContribuinteICMS(cpf_cnpj, colunaContribuinte, dbfC);
      except
        on E: Exception do
        begin
          contribuinteICMS := 9;
          LogMsg('log_erros_clientes.txt', 'Erro ao determinar contribuinte: ' + E.Message);
        end;
      end;


      try

        if cadastrado_em = 0 then
          cadastrado_em := Now;


        queryC.SQL.Text :=
          'INSERT INTO CLIENTE (' +
          'CODCLIENTE, NOME, CADASTRADO_EM, CIDADE, F_J, CPF_CNPJ, ' +
          'NIVEL_CREDITO, CPAIS, CONTRIBUINTEICMS, ALTERADO, DEBITO' +
          ') VALUES (' +
          ':CODIGO, :NOME, :CADASTRADO_EM, :CIDADE, :F_J, :CPF_CNPJ, ' +
          '5, :CPAIS, :CONTRIBUINTEICMS, :ALTERADO, 0)';


        queryC.ParamByName('CODIGO').AsInteger := codigo;
        queryC.ParamByName('NOME').AsString := nome;
        queryC.ParamByName('CADASTRADO_EM').AsDateTime := cadastrado_em;
        queryC.ParamByName('CIDADE').AsString := cidade;
        queryC.ParamByName('F_J').AsString := tipo_pessoa;
        queryC.ParamByName('CPF_CNPJ').AsString := cpf_cnpj;
        queryC.ParamByName('CPAIS').AsString := cpais;
        queryC.ParamByName('CONTRIBUINTEICMS').AsInteger := contribuinteICMS;
        queryC.ParamByName('ALTERADO').AsDateTime := cadastrado_em;

        queryC.ExecSQL;

        LogMsg('log_sucesso.txt', Format('Cliente %d inserido: %s - Data: %s',
          [codigo, nome, DateTimeToStr(cadastrado_em)]));

      except
        on E: Exception do
        begin
          LogMsg('log_erros_detalhados.txt',
            Format('ERRO cliente %d: %s - Detalhes: NOME="%s" F_J="%s" CPF_CNPJ="%s" CIDADE="%s" CADASTRADO_EM="%s" CONTRIBUINTE=%d - Erro: %s',
            [codigo, E.ClassName, nome, tipo_pessoa, cpf_cnpj, cidade,
             DateTimeToStr(cadastrado_em), contribuinteICMS, E.Message]));

          Inc(puladas);
          Dec(codigo);

          if transC.Active then
            transC.Rollback;


          try
            transC.StartTransaction;
          except
            on E2: Exception do
            begin
              LogMsg('log_erros_clientes.txt', 'Não foi possível reiniciar transação: ' + E2.Message);
              Break;
            end;
          end;
        end;
      end;


      if total > 0 then
        ProgressBar.Position := Round((atual / total) * 100);

      if (atual mod 50) = 0 then
      begin

        if transC.Active then
        begin
          try
            transC.CommitRetaining;
            LogMsg('log_progresso.txt', Format('CommitRetaining bem-sucedido no registro %d', [atual]));
          except
            on E: Exception do
            begin
              LogMsg('log_erros_clientes.txt', 'Erro no CommitRetaining: ' + E.Message);

              if transC.Active then
                transC.Rollback;
              transC.StartTransaction;
            end;
          end;
        end;

        LabelProgresso.Caption := Format('Importando %d/%d...', [atual, total]);
        Application.ProcessMessages;
      end;

      dbfC.Next;
    end;


    if transC.Active then
    begin
      try
        transC.Commit;
        LogMsg('log_progresso.txt', 'Commit final bem-sucedido');
      except
        on E: Exception do
        begin
          LogMsg('log_erros_clientes.txt', 'Erro no commit final: ' + E.Message);
          if transC.Active then
            transC.Rollback;
        end;
      end;
    end;

    ProgressBar.Position := 100;

    LabelProgresso.Caption := Format('✅ Clientes importados: %d | Pulados: %d | Config: %s',
      [codigo, puladas, ComboOpcaoContribuinte.Items[ComboOpcaoContribuinte.ItemIndex]]);
    ShowMessage(LabelProgresso.Caption);

  finally

    if dbfC.Active then
      dbfC.Close;

    FreeAndNil(dbfC);
    FreeAndNil(queryC);

    if transC.Active then
    begin
      try
        transC.Commit;
      except
        transC.Rollback;
      end;
    end;

    FreeAndNil(transC);

    if connC.Connected then
      connC.Close;

    FreeAndNil(connC);
  end;
end;

end.
