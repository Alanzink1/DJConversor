unit uImportarTributacao;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls;

type
  { TForm8 }
  TForm8 = class(TForm)
    ButtonImportarTributacao: TButton;
    ButtonBuscarDBF: TButton;
    ButtonBuscarFDB: TButton;
    ComboCFOP: TComboBox;
    EditCaminhoDBF: TEdit;
    EditCaminhoFDB: TEdit;
    LabelDBF: TLabel;
    LabelFDB: TLabel;
    LabelCFOP: TLabel;
    LabelProgresso: TLabel;
    OpenDialog: TOpenDialog;
    ProgressBar: TProgressBar;
    Panel1: TPanel;
    ComboCodigoFiscal: TComboBox;
    LabelCodigoFiscal: TLabel;
    CheckBoxImportarClassFiscal: TCheckBox;
    ComboICMS: TComboBox;
    LabelICMS: TLabel;
    ComboCSOSN: TComboBox;
    LabelCSOSN: TLabel;
    ComboFatorBase: TComboBox;
    LabelFatorBase: TLabel;
    ComboCSTPIS: TComboBox;
    LabelCSTPIS: TLabel;
    ComboDescricao: TComboBox;
    LabelDescricao: TLabel;
    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarTributacaoClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
  private
    procedure CarregarColunasDBF;
    procedure LogMsg(const nomeArq, msg: string);
    function SanitizeText(const S: string): string;
    function PrepareStringForDB(const S: string): string;
    procedure ImportarClassificacaoFiscal;
    procedure ImportarTributacaoICMS;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  Form8: TForm8;

implementation

{$R *.lfm}

{ TForm8 }

constructor TForm8.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TForm8.Destroy;
begin
  inherited Destroy;
end;

procedure TForm8.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor - Tributacao';
  Width := 700;
  Height := 550;
  Position := poScreenCenter;

  Panel1 := TPanel.Create(Self);
  Panel1.Parent := Self;
  Panel1.Align := alClient;
  Panel1.BevelOuter := bvNone;

  LabelDBF := TLabel.Create(Self);
  LabelDBF.Parent := Panel1;
  LabelDBF.Caption := 'Arquivo da Classificação Fiscal (.DBF):';
  LabelDBF.Left := 20;
  LabelDBF.Top := 20;

  EditCaminhoDBF := TEdit.Create(Self);
  EditCaminhoDBF.Parent := Panel1;
  EditCaminhoDBF.Left := 220;
  EditCaminhoDBF.Top := 16;
  EditCaminhoDBF.Width := 300;

  ButtonBuscarDBF := TButton.Create(Self);
  ButtonBuscarDBF.Parent := Panel1;
  ButtonBuscarDBF.Caption := '📂 Buscar';
  ButtonBuscarDBF.Left := 540;
  ButtonBuscarDBF.Top := 16;
  ButtonBuscarDBF.Width := 100;
  ButtonBuscarDBF.OnClick := @ButtonBuscarDBFClick;


  LabelFDB := TLabel.Create(Self);
  LabelFDB.Parent := Panel1;
  LabelFDB.Caption := 'Arquivo do banco (.FDB):';
  LabelFDB.Left := 20;
  LabelFDB.Top := 60;

  EditCaminhoFDB := TEdit.Create(Self);
  EditCaminhoFDB.Parent := Panel1;
  EditCaminhoFDB.Left := 220;
  EditCaminhoFDB.Top := 56;
  EditCaminhoFDB.Width := 300;

  ButtonBuscarFDB := TButton.Create(Self);
  ButtonBuscarFDB.Parent := Panel1;
  ButtonBuscarFDB.Caption := '📁 Buscar';
  ButtonBuscarFDB.Left := 540;
  ButtonBuscarFDB.Top := 56;
  ButtonBuscarFDB.Width := 100;
  ButtonBuscarFDB.OnClick := @ButtonBuscarFDBClick;


  LabelCodigoFiscal := TLabel.Create(Self);
  LabelCodigoFiscal.Parent := Panel1;
  LabelCodigoFiscal.Caption := 'Coluna Código Fiscal:';
  LabelCodigoFiscal.Left := 20;
  LabelCodigoFiscal.Top := 100;

  ComboCodigoFiscal := TComboBox.Create(Self);
  ComboCodigoFiscal.Parent := Panel1;
  ComboCodigoFiscal.Left := 220;
  ComboCodigoFiscal.Top := 96;
  ComboCodigoFiscal.Width := 300;


  LabelDescricao := TLabel.Create(Self);
  LabelDescricao.Parent := Panel1;
  LabelDescricao.Caption := 'Coluna Descrição:';
  LabelDescricao.Left := 20;
  LabelDescricao.Top := 140;

  ComboDescricao := TComboBox.Create(Self);
  ComboDescricao.Parent := Panel1;
  ComboDescricao.Left := 220;
  ComboDescricao.Top := 136;
  ComboDescricao.Width := 300;


  LabelCFOP := TLabel.Create(Self);
  LabelCFOP.Parent := Panel1;
  LabelCFOP.Caption := 'Coluna CFOP:';
  LabelCFOP.Left := 20;
  LabelCFOP.Top := 180;

  ComboCFOP := TComboBox.Create(Self);
  ComboCFOP.Parent := Panel1;
  ComboCFOP.Left := 220;
  ComboCFOP.Top := 176;
  ComboCFOP.Width := 300;


  LabelICMS := TLabel.Create(Self);
  LabelICMS.Parent := Panel1;
  LabelICMS.Caption := 'Coluna ICMS:';
  LabelICMS.Left := 20;
  LabelICMS.Top := 220;

  ComboICMS := TComboBox.Create(Self);
  ComboICMS.Parent := Panel1;
  ComboICMS.Left := 220;
  ComboICMS.Top := 216;
  ComboICMS.Width := 300;


  LabelCSOSN := TLabel.Create(Self);
  LabelCSOSN.Parent := Panel1;
  LabelCSOSN.Caption := 'Coluna CSOSN:';
  LabelCSOSN.Left := 20;
  LabelCSOSN.Top := 260;

  ComboCSOSN := TComboBox.Create(Self);
  ComboCSOSN.Parent := Panel1;
  ComboCSOSN.Left := 220;
  ComboCSOSN.Top := 256;
  ComboCSOSN.Width := 300;


  LabelFatorBase := TLabel.Create(Self);
  LabelFatorBase.Parent := Panel1;
  LabelFatorBase.Caption := 'Coluna Fator Base:';
  LabelFatorBase.Left := 20;
  LabelFatorBase.Top := 300;

  ComboFatorBase := TComboBox.Create(Self);
  ComboFatorBase.Parent := Panel1;
  ComboFatorBase.Left := 220;
  ComboFatorBase.Top := 296;
  ComboFatorBase.Width := 300;


  LabelCSTPIS := TLabel.Create(Self);
  LabelCSTPIS.Parent := Panel1;
  LabelCSTPIS.Caption := 'Coluna CST PIS:';
  LabelCSTPIS.Left := 20;
  LabelCSTPIS.Top := 340;

  ComboCSTPIS := TComboBox.Create(Self);
  ComboCSTPIS.Parent := Panel1;
  ComboCSTPIS.Left := 220;
  ComboCSTPIS.Top := 336;
  ComboCSTPIS.Width := 300;


  CheckBoxImportarClassFiscal := TCheckBox.Create(Self);
  CheckBoxImportarClassFiscal.Parent := Panel1;
  CheckBoxImportarClassFiscal.Caption := 'Importar Classificação Fiscal e Tributação ICMS';
  CheckBoxImportarClassFiscal.Left := 20;
  CheckBoxImportarClassFiscal.Top := 380;
  CheckBoxImportarClassFiscal.Width := 300;
  CheckBoxImportarClassFiscal.Checked := True;


  ButtonImportarTributacao := TButton.Create(Self);
  ButtonImportarTributacao.Parent := Panel1;
  ButtonImportarTributacao.Caption := '🚀 Importar Tributação Completa';
  ButtonImportarTributacao.Left := 20;
  ButtonImportarTributacao.Top := 420;
  ButtonImportarTributacao.Width := 250;
  ButtonImportarTributacao.Height := 40;
  ButtonImportarTributacao.OnClick := @ButtonImportarTributacaoClick;


  ProgressBar := TProgressBar.Create(Self);
  ProgressBar.Parent := Panel1;
  ProgressBar.Left := 20;
  ProgressBar.Top := 480;
  ProgressBar.Width := 620;
  ProgressBar.Min := 0;
  ProgressBar.Max := 100;

  LabelProgresso := TLabel.Create(Self);
  LabelProgresso.Parent := Panel1;
  LabelProgresso.Left := 20;
  LabelProgresso.Top := 510;
  LabelProgresso.Caption := '';

  OpenDialog := TOpenDialog.Create(Self);
end;

procedure TForm8.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
  Form8 := nil;
end;

procedure TForm8.ButtonBuscarDBFClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Arquivos DBF|*.dbf';
  if OpenDialog.Execute then
  begin
    EditCaminhoDBF.Text := OpenDialog.FileName;
    CarregarColunasDBF;
  end;
end;

procedure TForm8.ButtonBuscarFDBClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Banco Firebird|*.fdb';
  if OpenDialog.Execute then
    EditCaminhoFDB.Text := OpenDialog.FileName;
end;

procedure TForm8.ButtonImportarTributacaoClick(Sender: TObject);
begin
  if (EditCaminhoDBF.Text = '') or (EditCaminhoFDB.Text = '') then
  begin
    ShowMessage('Selecione os arquivos DBF e FDB antes de importar.');
    Exit;
  end;

  if ComboCodigoFiscal.Text = '' then
  begin
    ShowMessage('Selecione a coluna Código Fiscal para importar.');
    Exit;
  end;


  if CheckBoxImportarClassFiscal.Checked then
  begin
    ImportarClassificacaoFiscal;
    ImportarTributacaoICMS;
  end
  else
    ShowMessage('Nenhuma operação selecionada.');
end;

procedure TForm8.ImportarTributacaoICMS;
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfFile: TDbf;
  configsUnicas: TStringList;
  codFiscal, descricao, cfop, csosn, cst: string;
  aliquotaICMS, baseICMS: Double;
  total, atual, novoCount, lastID_ICMS: Integer;
begin
  LabelProgresso.Caption := 'Importando tributação ICMS...';
  ProgressBar.Position := 0;
  Application.ProcessMessages;

  configsUnicas := TStringList.Create;
  configsUnicas.Sorted := True;
  configsUnicas.Duplicates := dupIgnore;

  conn := TIBConnection.Create(nil);
  trans := TSQLTransaction.Create(nil);
  query := TSQLQuery.Create(nil);
  dbfFile := TDbf.Create(nil);

  try
    try

      conn.DatabaseName := EditCaminhoFDB.Text;
      conn.UserName := 'sysdba';
      conn.Password := 'masterkey';
      conn.CharSet := 'UTF8';
      conn.Params.Add('lc_ctype=UTF8');
      conn.Params.Add('sql_dialect=3');
      conn.Transaction := trans;

      trans.DataBase := conn;
      query.DataBase := conn;
      query.Transaction := trans;

      conn.Open;


      query.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE UPPER(RDB$RELATION_NAME) = ''ICMS''';
      query.Open;
      if query.IsEmpty then
      begin
        query.Close;
        ShowMessage('Tabela ICMS não encontrada no banco de dados.');
        Exit;
      end;
      query.Close;


      query.SQL.Text := 'SELECT COALESCE(MAX(ID_ICMS), 0) AS N FROM ICMS';
      query.Open;
      lastID_ICMS := StrToIntDef(Trim(query.Fields[0].AsString), 0);
      query.Close;


      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D FROM ICMS';
      query.Open;
      while not query.EOF do
      begin
        configsUnicas.Add(query.FieldByName('D').AsString);
        query.Next;
      end;
      query.Close;


      dbfFile.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
      dbfFile.TableName := ExtractFileName(EditCaminhoDBF.Text);
      dbfFile.Open;

      total := dbfFile.RecordCount;
      atual := 0;
      novoCount := 0;

      if trans.Active then
      trans.Commit;

      trans.StartTransaction;

      dbfFile.First;
      while not dbfFile.EOF do
      begin
        Inc(atual);


        codFiscal := Trim(dbfFile.FieldByName(ComboCodigoFiscal.Text).AsString);

        if codFiscal <> '' then
        begin

          if ComboDescricao.Text <> '' then
            descricao := 'CF ' + codFiscal + ' - ' + Trim(dbfFile.FieldByName(ComboDescricao.Text).AsString)
          else
            descricao := 'CF ' + codFiscal;


          if configsUnicas.IndexOf(UpperCase(descricao)) = -1 then
          begin
            Inc(lastID_ICMS);


            cfop := '';
            if ComboCFOP.Text <> '' then
              cfop := Trim(dbfFile.FieldByName(ComboCFOP.Text).AsString);


            csosn := '';
            if ComboCSOSN.Text <> '' then
              csosn := Trim(dbfFile.FieldByName(ComboCSOSN.Text).AsString);


            cst := '00';
            if csosn <> '' then
              cst := Copy(csosn, 1, 2)
            else if ComboCSTPIS.Text <> '' then
              cst := Trim(dbfFile.FieldByName(ComboCSTPIS.Text).AsString);


            aliquotaICMS := 0;
            if ComboICMS.Text <> '' then
              aliquotaICMS := dbfFile.FieldByName(ComboICMS.Text).AsFloat;


            baseICMS := 100.0;
            if ComboFatorBase.Text <> '' then
              baseICMS := dbfFile.FieldByName(ComboFatorBase.Text).AsFloat * 100;


            if query.Active then
              query.Close;

            query.SQL.Text :=
              'INSERT INTO ICMS (' +
              'ID_ICMS, DESCRICAO, CST, MODBC, BASEICMS, ALIQICMS, ' +
              'CSOSN, CFOP, ORIGEM) ' +
              'VALUES (' +
              ':ID_ICMS, :DESCRICAO, :CST, :MODBC, :BASEICMS, :ALIQICMS, ' +
              ':CSOSN, :CFOP, :ORIGEM)';

            query.ParamByName('ID_ICMS').AsInteger := lastID_ICMS;
            query.ParamByName('DESCRICAO').AsString := Copy(descricao, 1, 100);
            query.ParamByName('CST').AsString := Copy(cst, 1, 2);
            query.ParamByName('MODBC').AsInteger := 3;
            query.ParamByName('BASEICMS').AsFloat := baseICMS;
            query.ParamByName('ALIQICMS').AsFloat := aliquotaICMS;
            query.ParamByName('CSOSN').AsString := Copy(csosn, 1, 4);
            query.ParamByName('CFOP').AsString := Copy(cfop, 1, 10);
            query.ParamByName('ORIGEM').AsInteger := 0;

            query.ExecSQL;

            configsUnicas.Add(UpperCase(descricao));
            Inc(novoCount);
          end;
        end;


        if (atual mod 10) = 0 then
        begin
          ProgressBar.Position := Round((atual / total) * 100);
          LabelProgresso.Caption := Format('ICMS: %d/%d - Novos: %d', [atual, total, novoCount]);
          Application.ProcessMessages;
        end;

        dbfFile.Next;
      end;

      trans.Commit;

      LabelProgresso.Caption := Format('✅ Tributação ICMS: %d novos de %d registros', [novoCount, total]);
      ShowMessage(LabelProgresso.Caption);

    except
      on E: Exception do
      begin
        if trans.Active then
          trans.Rollback;
        ShowMessage('Erro ao importar tributação ICMS: ' + E.Message);
        LogMsg('log_erros_icms.txt', 'Erro importar ICMS: ' + E.Message);
      end;
    end;

  finally
    if query.Active then
      query.Close;
    dbfFile.Free;
    query.Free;
    trans.Free;
    conn.Free;
    configsUnicas.Free;
  end;
end;


procedure TForm8.ImportarClassificacaoFiscal;
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfFile: TDbf;
  classFiscaisUnicas: TStringList;
  codFiscal, descFiscal, cfop, cst: string;
  aliquotaICMS, aliquotaIPI: Double;
  total, atual, novoCount: Integer;
  tabelaExiste: Boolean;
begin
  if ComboCodigoFiscal.Text = '' then
    Exit;

  LabelProgresso.Caption := 'Importando classificação fiscal...';
  ProgressBar.Position := 0;
  Application.ProcessMessages;

  classFiscaisUnicas := TStringList.Create;
  classFiscaisUnicas.Sorted := True;
  classFiscaisUnicas.Duplicates := dupIgnore;

  conn := TIBConnection.Create(nil);
  trans := TSQLTransaction.Create(nil);
  query := TSQLQuery.Create(nil);
  dbfFile := TDbf.Create(nil);

  try
    try

      conn.DatabaseName := EditCaminhoFDB.Text;
      conn.UserName := 'sysdba';
      conn.Password := 'masterkey';
      conn.CharSet := 'UTF8';
      conn.Params.Add('lc_ctype=UTF8');
      conn.Params.Add('sql_dialect=3');
      conn.Transaction := trans;

      trans.DataBase := conn;
      query.DataBase := conn;
      query.Transaction := trans;

      conn.Open;


      tabelaExiste := False;
      try
        if query.Active then query.Close;
        query.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE RDB$RELATION_NAME = ''CLASSIFICACAO_FISCAL''';
        query.Open;
        tabelaExiste := not query.IsEmpty;
        query.Close;
      except
        tabelaExiste := False;
      end;


      if not tabelaExiste then
      begin
        if query.Active then query.Close;


        if trans.Active then
          trans.Commit;


        query.SQL.Text :=
          'CREATE TABLE CLASSIFICACAO_FISCAL (' +
          'CODIGO VARCHAR(10) NOT NULL PRIMARY KEY, ' +
          'DESCRICAO VARCHAR(100), ' +
          'ALIQUOTA_ICMS NUMERIC(15,4), ' +
          'ALIQUOTA_IPI NUMERIC(15,4), ' +
          'CFOP VARCHAR(10), ' +
          'CST VARCHAR(4))';
        query.ExecSQL;


        trans.Commit;


        ShowMessage('Tabela CLASSIFICACAO_FISCAL criada com sucesso!');
      end;

      if trans.Active then
         trans.Commit;


      trans.StartTransaction;


      if query.Active then query.Close;
      query.SQL.Text := 'SELECT UPPER(CODIGO) AS C FROM CLASSIFICACAO_FISCAL';
      query.Open;
      while not query.EOF do
      begin
        classFiscaisUnicas.Add(query.FieldByName('C').AsString);
        query.Next;
      end;
      query.Close;


      dbfFile.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
      dbfFile.TableName := ExtractFileName(EditCaminhoDBF.Text);
      dbfFile.Open;

      total := dbfFile.RecordCount;
      atual := 0;
      novoCount := 0;


      if not trans.Active then
        trans.StartTransaction;

      dbfFile.First;
      while not dbfFile.EOF do
      begin
        Inc(atual);


        codFiscal := Trim(dbfFile.FieldByName(ComboCodigoFiscal.Text).AsString);

        if (codFiscal <> '') and (classFiscaisUnicas.IndexOf(UpperCase(codFiscal)) = -1) then
        begin

          if ComboDescricao.Text <> '' then
            descFiscal := Trim(dbfFile.FieldByName(ComboDescricao.Text).AsString)
          else
            descFiscal := 'Código Fiscal ' + codFiscal;


          cfop := '';
          if ComboCFOP.Text <> '' then
            cfop := Trim(dbfFile.FieldByName(ComboCFOP.Text).AsString);


          aliquotaICMS := 0;
          aliquotaIPI := 0;
          if ComboICMS.Text <> '' then
            aliquotaICMS := dbfFile.FieldByName(ComboICMS.Text).AsFloat;


          cst := '000';
          if ComboCSOSN.Text <> '' then
            cst := Trim(dbfFile.FieldByName(ComboCSOSN.Text).AsString)
          else if ComboCSTPIS.Text <> '' then
            cst := Trim(dbfFile.FieldByName(ComboCSTPIS.Text).AsString);


          if query.Active then
            query.Close;

          query.SQL.Text :=
            'INSERT INTO CLASSIFICACAO_FISCAL (CODIGO, DESCRICAO, ALIQUOTA_ICMS, ALIQUOTA_IPI, CFOP, CST) ' +
            'VALUES (:COD, :DESC, :ICMS, :IPI, :CFOP, :CST)';
          query.ParamByName('COD').AsString := Copy(codFiscal, 1, 10);
          query.ParamByName('DESC').AsString := Copy(descFiscal, 1, 100);
          query.ParamByName('ICMS').AsFloat := aliquotaICMS;
          query.ParamByName('IPI').AsFloat := aliquotaIPI;
          query.ParamByName('CFOP').AsString := Copy(cfop, 1, 10);
          query.ParamByName('CST').AsString := Copy(cst, 1, 4);
          query.ExecSQL;

          classFiscaisUnicas.Add(UpperCase(codFiscal));
          Inc(novoCount);
        end;


        if (atual mod 10) = 0 then
        begin
          ProgressBar.Position := Round((atual / total) * 100);
          LabelProgresso.Caption := Format('Class. Fiscal: %d/%d - Novas: %d', [atual, total, novoCount]);
          Application.ProcessMessages;
        end;

        dbfFile.Next;
      end;

      trans.Commit;

      LabelProgresso.Caption := Format('✅ Classificação fiscal: %d novas de %d registros', [novoCount, total]);

    except
      on E: Exception do
      begin
        if trans.Active then
          trans.Rollback;
        ShowMessage('Erro ao importar classificação fiscal: ' + E.Message);
        LogMsg('log_erros_classfiscal.txt', 'Erro importar class. fiscal: ' + E.Message);
      end;
    end;

  finally
    if query.Active then
      query.Close;
    dbfFile.Free;
    query.Free;
    trans.Free;
    conn.Free;
    classFiscaisUnicas.Free;
  end;
end;



procedure TForm8.CarregarColunasDBF;
var
  dbfFile: TDbf;
  i: Integer;
begin
  if not FileExists(EditCaminhoDBF.Text) then
    Exit;

  dbfFile := TDbf.Create(nil);
  try
    dbfFile.FilePathFull := ExtractFilePath(EditCaminhoDBF.Text);
    dbfFile.TableName := ExtractFileName(EditCaminhoDBF.Text);
    dbfFile.Open;


    ComboCodigoFiscal.Clear;
    ComboDescricao.Clear;
    ComboCFOP.Clear;
    ComboICMS.Clear;
    ComboCSOSN.Clear;
    ComboFatorBase.Clear;
    ComboCSTPIS.Clear;


    for i := 0 to dbfFile.FieldDefs.Count - 1 do
    begin
      ComboCodigoFiscal.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboDescricao.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboCFOP.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboICMS.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboCSOSN.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboFatorBase.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboCSTPIS.Items.Add(dbfFile.FieldDefs[i].Name);
    end;


    if ComboCodigoFiscal.Items.IndexOf('CODCFI') >= 0 then
      ComboCodigoFiscal.Text := 'CODCFI'
    else if ComboCodigoFiscal.Items.IndexOf('CODIGO') >= 0 then
      ComboCodigoFiscal.Text := 'CODIGO';

    if ComboDescricao.Items.IndexOf('DESCRICAO') >= 0 then
      ComboDescricao.Text := 'DESCRICAO';

    if ComboCFOP.Items.IndexOf('CFOP') >= 0 then
      ComboCFOP.Text := 'CFOP';

    if ComboICMS.Items.IndexOf('ICMS') >= 0 then
      ComboICMS.Text := 'ICMS';

    if ComboCSOSN.Items.IndexOf('CSOSN') >= 0 then
      ComboCSOSN.Text := 'CSOSN';

    if ComboFatorBase.Items.IndexOf('FATOR_BASE') >= 0 then
      ComboFatorBase.Text := 'FATOR_BASE';

    if ComboCSTPIS.Items.IndexOf('CSTPIS') >= 0 then
      ComboCSTPIS.Text := 'CSTPIS';

    dbfFile.Close;
    ShowMessage('Colunas carregadas com sucesso! Selecione manualmente as colunas não detectadas automaticamente.');
  finally
    dbfFile.Free;
  end;
end;

procedure TForm8.LogMsg(const nomeArq, msg: string);
var
  f: TextFile;
  caminho: string;
begin
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

function TForm8.SanitizeText(const S: string): string;
var
  i: Integer;
  tmp: string;
begin
  tmp := UpperCase(Trim(S));
  Result := '';
  for i := 1 to Length(tmp) do
  begin
    if CharInSet(tmp[i], ['A'..'Z', '0'..'9', ' ', '-', '_', '/', '.', ',', '(', ')', ':']) then
      Result := Result + tmp[i];
  end;
  Result := Trim(StringReplace(Result, '  ', ' ', [rfReplaceAll]));
end;

function TForm8.PrepareStringForDB(const S: string): string;
var
  i: Integer;
  temp: string;
begin
  temp := Trim(S);
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

end.
