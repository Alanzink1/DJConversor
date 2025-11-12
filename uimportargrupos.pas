unit uImportarGrupos;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls;

type
  { TForm7 }
  TForm7 = class(TForm)
    ButtonImportarGrupos: TButton;
    ButtonBuscarDBF: TButton;
    ButtonBuscarFDB: TButton;
    ComboDescricao: TComboBox;
    ComboCodGrupo: TComboBox;
    EditCaminhoDBF: TEdit;
    EditCaminhoFDB: TEdit;
    LabelDBF: TLabel;
    LabelFDB: TLabel;
    LabelDescricao: TLabel;
    LabelCodGrupo: TLabel;
    LabelProgresso: TLabel;
    OpenDialog: TOpenDialog;
    ProgressBar: TProgressBar;
    Panel1: TPanel;
    CheckUsarCodGrupo: TCheckBox;
    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarGruposClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure CheckUsarCodGrupoClick(Sender: TObject);
  private
    procedure CarregarColunasDBF;
    procedure LogMsg(const nomeArq, msg: string);
    function SanitizeText(const S: string): string;
    function PrepareStringForDB(const S: string): string;
    function TableExists(Q: TSQLQuery; const tableName: string): Boolean;
  public
    constructor Create(AOwner: TComponent); override;
    destructor Destroy; override;
  end;

var
  Form7: TForm7;

implementation

{$R *.lfm}

{ TForm7 }

constructor TForm7.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TForm7.Destroy;
begin
  inherited Destroy;
end;

procedure TForm7.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor - Grupos';
  Width := 600;
  Height := 350;
  Position := poScreenCenter;


  Panel1 := TPanel.Create(Self);
  Panel1.Parent := Self;
  Panel1.Align := alClient;
  Panel1.BevelOuter := bvNone;


  LabelDBF := TLabel.Create(Self);
  LabelDBF.Parent := Panel1;
  LabelDBF.Caption := 'Arquivo da planilha (.DBF):';
  LabelDBF.Left := 20;
  LabelDBF.Top := 20;

  EditCaminhoDBF := TEdit.Create(Self);
  EditCaminhoDBF.Parent := Panel1;
  EditCaminhoDBF.Left := 200;
  EditCaminhoDBF.Top := 16;
  EditCaminhoDBF.Width := 250;

  ButtonBuscarDBF := TButton.Create(Self);
  ButtonBuscarDBF.Parent := Panel1;
  ButtonBuscarDBF.Caption := '📂 Buscar';
  ButtonBuscarDBF.Left := 460;
  ButtonBuscarDBF.Top := 16;
  ButtonBuscarDBF.Width := 80;
  ButtonBuscarDBF.OnClick := @ButtonBuscarDBFClick;


  LabelFDB := TLabel.Create(Self);
  LabelFDB.Parent := Panel1;
  LabelFDB.Caption := 'Arquivo do banco (.FDB):';
  LabelFDB.Left := 20;
  LabelFDB.Top := 60;

  EditCaminhoFDB := TEdit.Create(Self);
  EditCaminhoFDB.Parent := Panel1;
  EditCaminhoFDB.Left := 200;
  EditCaminhoFDB.Top := 56;
  EditCaminhoFDB.Width := 250;

  ButtonBuscarFDB := TButton.Create(Self);
  ButtonBuscarFDB.Parent := Panel1;
  ButtonBuscarFDB.Caption := '📁 Buscar';
  ButtonBuscarFDB.Left := 460;
  ButtonBuscarFDB.Top := 56;
  ButtonBuscarFDB.Width := 80;
  ButtonBuscarFDB.OnClick := @ButtonBuscarFDBClick;


  LabelDescricao := TLabel.Create(Self);
  LabelDescricao.Parent := Panel1;
  LabelDescricao.Caption := 'Coluna Descrição:';
  LabelDescricao.Left := 20;
  LabelDescricao.Top := 100;

  ComboDescricao := TComboBox.Create(Self);
  ComboDescricao.Parent := Panel1;
  ComboDescricao.Left := 200;
  ComboDescricao.Top := 96;
  ComboDescricao.Width := 250;


  CheckUsarCodGrupo := TCheckBox.Create(Self);
  CheckUsarCodGrupo.Parent := Panel1;
  CheckUsarCodGrupo.Caption := 'Usar código do grupo existente (opcional)';
  CheckUsarCodGrupo.Left := 20;
  CheckUsarCodGrupo.Top := 135;
  CheckUsarCodGrupo.Width := 250;
  CheckUsarCodGrupo.OnClick := @CheckUsarCodGrupoClick;


  LabelCodGrupo := TLabel.Create(Self);
  LabelCodGrupo.Parent := Panel1;
  LabelCodGrupo.Caption := 'Coluna Código:';
  LabelCodGrupo.Left := 20;
  LabelCodGrupo.Top := 165;
  LabelCodGrupo.Enabled := False;

  ComboCodGrupo := TComboBox.Create(Self);
  ComboCodGrupo.Parent := Panel1;
  ComboCodGrupo.Left := 200;
  ComboCodGrupo.Top := 161;
  ComboCodGrupo.Width := 250;
  ComboCodGrupo.Enabled := False;
  ComboCodGrupo.Style := csDropDownList;


  ButtonImportarGrupos := TButton.Create(Self);
  ButtonImportarGrupos.Parent := Panel1;
  ButtonImportarGrupos.Caption := '🚀 Importar Grupos';
  ButtonImportarGrupos.Left := 20;
  ButtonImportarGrupos.Top := 200;
  ButtonImportarGrupos.Width := 200;
  ButtonImportarGrupos.Height := 40;
  ButtonImportarGrupos.OnClick := @ButtonImportarGruposClick;


  ProgressBar := TProgressBar.Create(Self);
  ProgressBar.Parent := Panel1;
  ProgressBar.Left := 20;
  ProgressBar.Top := 250;
  ProgressBar.Width := 520;
  ProgressBar.Min := 0;
  ProgressBar.Max := 100;

  LabelProgresso := TLabel.Create(Self);
  LabelProgresso.Parent := Panel1;
  LabelProgresso.Left := 20;
  LabelProgresso.Top := 280;
  LabelProgresso.Caption := '';

  OpenDialog := TOpenDialog.Create(Self);
end;

procedure TForm7.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
  Form7 := nil;
end;

procedure TForm7.CheckUsarCodGrupoClick(Sender: TObject);
begin

  LabelCodGrupo.Enabled := CheckUsarCodGrupo.Checked;
  ComboCodGrupo.Enabled := CheckUsarCodGrupo.Checked;

  if CheckUsarCodGrupo.Checked and (ComboCodGrupo.Items.Count = 0) then
  begin

    if FileExists(EditCaminhoDBF.Text) then
      CarregarColunasDBF;
  end;
end;

procedure TForm7.ButtonBuscarDBFClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Arquivos DBF|*.dbf';
  if OpenDialog.Execute then
  begin
    EditCaminhoDBF.Text := OpenDialog.FileName;
    CarregarColunasDBF;
  end;
end;

procedure TForm7.ButtonBuscarFDBClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Banco Firebird|*.fdb';
  if OpenDialog.Execute then
    EditCaminhoFDB.Text := OpenDialog.FileName;
end;

procedure TForm7.ButtonImportarGruposClick(Sender: TObject);
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfFile: TDbf;
  gruposUnicos: TStringList;
  descGrupo, codGrupoStr: string;
  total, atual, novoCount: Integer;
  lastCodGrupo, codGrupo: Integer;
  usarCodGrupo: Boolean;
begin
  if (EditCaminhoDBF.Text = '') or (EditCaminhoFDB.Text = '') then
  begin
    ShowMessage('Selecione os arquivos DBF e FDB antes de importar.');
    Exit;
  end;

  if ComboDescricao.Text = '' then
  begin
    ShowMessage('Selecione a coluna Descrição para importar.');
    Exit;
  end;

  usarCodGrupo := CheckUsarCodGrupo.Checked and (ComboCodGrupo.Text <> '');

  LabelProgresso.Caption := 'Importando grupos...';
  ProgressBar.Position := 0;
  Application.ProcessMessages;

  gruposUnicos := TStringList.Create;
  gruposUnicos.Sorted := True;
  gruposUnicos.Duplicates := dupIgnore;

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


      if not TableExists(query, 'GRUPO') then
      begin
        ShowMessage('Tabela GRUPO não encontrada no banco de dados. Criando tabela...');


        if trans.Active then
          trans.Commit;

        query.SQL.Text := 'CREATE TABLE GRUPO (CODGRUPO INTEGER PRIMARY KEY, DESCRICAO VARCHAR(50))';
        query.ExecSQL;
        trans.Commit;
      end;


      query.SQL.Text := 'SELECT COALESCE(MAX(CODGRUPO), 0) AS N FROM GRUPO';
      query.Open;
      lastCodGrupo := StrToIntDef(Trim(query.Fields[0].AsString), 0);
      query.Close;


      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D FROM GRUPO';
      query.Open;
      while not query.EOF do
      begin
        gruposUnicos.Add(query.FieldByName('D').AsString);
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
      begin
        trans.Commit;
      end;


      trans.StartTransaction;

      dbfFile.First;
      while not dbfFile.EOF do
      begin
        Inc(atual);
        descGrupo := Trim(dbfFile.FieldByName(ComboDescricao.Text).AsString);


        if usarCodGrupo then
        begin
          codGrupoStr := Trim(dbfFile.FieldByName(ComboCodGrupo.Text).AsString);
          codGrupo := StrToIntDef(codGrupoStr, 0);
        end
        else
        begin
          codGrupo := 0;
        end;

        if (descGrupo <> '') and (gruposUnicos.IndexOf(UpperCase(descGrupo)) = -1) then
        begin

          if usarCodGrupo and (codGrupo > 0) then
          begin

            query.Close;
            query.SQL.Text := 'SELECT 1 FROM GRUPO WHERE CODGRUPO = :COD';
            query.ParamByName('COD').AsInteger := codGrupo;
            query.Open;

            if not query.IsEmpty then
            begin

              LogMsg('log_erros_grupos.txt',
                Format('Código de grupo já existe: %d - %s. Pulando registro.', [codGrupo, descGrupo]));
              query.Close;
              dbfFile.Next;
              Continue;
            end;
            query.Close;


            lastCodGrupo := codGrupo;
          end
          else
          begin

            Inc(lastCodGrupo);
          end;

          query.SQL.Text := 'INSERT INTO GRUPO (CODGRUPO, DESCRICAO) VALUES (:COD, :DESC)';
          query.ParamByName('COD').AsInteger := lastCodGrupo;
          query.ParamByName('DESC').AsString := Copy(descGrupo, 1, 50);

          try
            query.ExecSQL;
            gruposUnicos.Add(UpperCase(descGrupo));
            Inc(novoCount);

            if usarCodGrupo and (codGrupo > 0) then
              LogMsg('log_grupos_importados.txt',
                Format('Grupo importado com código específico: %d - %s', [codGrupo, descGrupo]))
            else
              LogMsg('log_grupos_importados.txt',
                Format('Grupo importado com código automático: %d - %s', [lastCodGrupo, descGrupo]));

          except
            on E: Exception do
            begin
              LogMsg('log_erros_grupos.txt',
                Format('Erro ao inserir grupo: %s - %s', [descGrupo, E.Message]));
            end;
          end;
        end;


        if (atual mod 100) = 0 then
        begin
          ProgressBar.Position := Round((atual / total) * 100);
          if usarCodGrupo then
            LabelProgresso.Caption := Format('Processando: %d/%d - Novos: %d (com código)', [atual, total, novoCount])
          else
            LabelProgresso.Caption := Format('Processando: %d/%d - Novos: %d', [atual, total, novoCount]);
          Application.ProcessMessages;


          trans.Commit;


          if not trans.Active then
            trans.StartTransaction;
        end;

        dbfFile.Next;
      end;


      if trans.Active then
        trans.Commit;

      ProgressBar.Position := 100;

      if usarCodGrupo then
        LabelProgresso.Caption := Format('✅ Grupos importados: %d novos de %d registros (com código específico)', [novoCount, total])
      else
        LabelProgresso.Caption := Format('✅ Grupos importados: %d novos de %d registros', [novoCount, total]);


      if novoCount > 0 then
      begin
        if usarCodGrupo then
          ShowMessage(Format('Importação concluída!%sForam importados %d novos grupos com códigos específicos.', [sLineBreak, novoCount]))
        else
          ShowMessage(Format('Importação concluída!%sForam importados %d novos grupos.', [sLineBreak, novoCount]));
      end
      else
        ShowMessage('Nenhum novo grupo foi importado. Todos os grupos já existiam no banco.');

    except
      on E: Exception do
      begin
        if trans.Active then
        begin
          trans.Rollback;
        end;
        ShowMessage('Erro ao importar grupos: ' + E.Message);
        LogMsg('log_erros_grupos.txt', 'Erro geral importar grupos: ' + E.Message);
      end;
    end;

  finally
    if Assigned(dbfFile) then
    begin
      if dbfFile.Active then
        dbfFile.Close;
      FreeAndNil(dbfFile);
    end;

    if Assigned(query) then
      FreeAndNil(query);

    if Assigned(trans) then
      FreeAndNil(trans);

    if Assigned(conn) then
    begin
      if conn.Connected then
        conn.Close;
      FreeAndNil(conn);
    end;

    if Assigned(gruposUnicos) then
      FreeAndNil(gruposUnicos);
  end;
end;

procedure TForm7.CarregarColunasDBF;
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

    ComboDescricao.Clear;
    ComboCodGrupo.Clear;


    for i := 0 to dbfFile.FieldDefs.Count - 1 do
    begin
      ComboDescricao.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboCodGrupo.Items.Add(dbfFile.FieldDefs[i].Name);
    end;

    dbfFile.Close;


    if ComboDescricao.Items.Count > 0 then
      ComboDescricao.ItemIndex := 0;

    if ComboCodGrupo.Items.Count > 0 then
      ComboCodGrupo.ItemIndex := 0;

    ShowMessage('Colunas carregadas com sucesso!');
  finally
    dbfFile.Free;
  end;
end;

procedure TForm7.LogMsg(const nomeArq, msg: string);
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

function TForm7.SanitizeText(const S: string): string;
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

function TForm7.PrepareStringForDB(const S: string): string;
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

function TForm7.TableExists(Q: TSQLQuery; const tableName: string): Boolean;
begin
  try
    Q.Close;
    Q.SQL.Text := 'SELECT 1 FROM RDB$RELATIONS WHERE UPPER(RDB$RELATION_NAME) = :T';
    Q.ParamByName('T').AsString := UpperCase(tableName);
    Q.Open;
    Result := not Q.IsEmpty;
    Q.Close;
  except
    Result := False;
  end;
end;

end.
