unit uImportarMarcas;

{$mode ObjFPC}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls;

type
  { TForm6 }
  TForm6 = class(TForm)
    ButtonImportarMarcas: TButton;
    ButtonBuscarDBF: TButton;
    ButtonBuscarFDB: TButton;
    ComboDescricao: TComboBox;
    ComboCodMarca: TComboBox;
    EditCaminhoDBF: TEdit;
    EditCaminhoFDB: TEdit;
    LabelDBF: TLabel;
    LabelFDB: TLabel;
    LabelDescricao: TLabel;
    LabelCodMarca: TLabel;
    LabelProgresso: TLabel;
    OpenDialog: TOpenDialog;
    ProgressBar: TProgressBar;
    Panel1: TPanel;
    CheckUsarCodMarca: TCheckBox;
    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarMarcasClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure CheckUsarCodMarcaClick(Sender: TObject);
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
  Form6: TForm6;

implementation

{$R *.lfm}

{ TForm6 }

constructor TForm6.Create(AOwner: TComponent);
begin
  inherited Create(AOwner);
end;

destructor TForm6.Destroy;
begin
  inherited Destroy;
end;

procedure TForm6.FormCreate(Sender: TObject);
begin
  Caption := 'DJConversor - Marcas';
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


  CheckUsarCodMarca := TCheckBox.Create(Self);
  CheckUsarCodMarca.Parent := Panel1;
  CheckUsarCodMarca.Caption := 'Usar código da marca existente (opcional)';
  CheckUsarCodMarca.Left := 20;
  CheckUsarCodMarca.Top := 135;
  CheckUsarCodMarca.Width := 250;
  CheckUsarCodMarca.OnClick := @CheckUsarCodMarcaClick;


  LabelCodMarca := TLabel.Create(Self);
  LabelCodMarca.Parent := Panel1;
  LabelCodMarca.Caption := 'Coluna Código:';
  LabelCodMarca.Left := 20;
  LabelCodMarca.Top := 165;
  LabelCodMarca.Enabled := False;

  ComboCodMarca := TComboBox.Create(Self);
  ComboCodMarca.Parent := Panel1;
  ComboCodMarca.Left := 200;
  ComboCodMarca.Top := 161;
  ComboCodMarca.Width := 250;
  ComboCodMarca.Enabled := False;
  ComboCodMarca.Style := csDropDownList;


  ButtonImportarMarcas := TButton.Create(Self);
  ButtonImportarMarcas.Parent := Panel1;
  ButtonImportarMarcas.Caption := '🚀 Importar Marcas';
  ButtonImportarMarcas.Left := 20;
  ButtonImportarMarcas.Top := 200;
  ButtonImportarMarcas.Width := 200;
  ButtonImportarMarcas.Height := 40;
  ButtonImportarMarcas.OnClick := @ButtonImportarMarcasClick;


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

procedure TForm6.FormClose(Sender: TObject; var CloseAction: TCloseAction);
begin
  CloseAction := caFree;
  Form6 := nil;
end;

procedure TForm6.CheckUsarCodMarcaClick(Sender: TObject);
begin

  LabelCodMarca.Enabled := CheckUsarCodMarca.Checked;
  ComboCodMarca.Enabled := CheckUsarCodMarca.Checked;

  if CheckUsarCodMarca.Checked and (ComboCodMarca.Items.Count = 0) then
  begin

    if FileExists(EditCaminhoDBF.Text) then
      CarregarColunasDBF;
  end;
end;

procedure TForm6.ButtonBuscarDBFClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Arquivos DBF|*.dbf';
  if OpenDialog.Execute then
  begin
    EditCaminhoDBF.Text := OpenDialog.FileName;
    CarregarColunasDBF;
  end;
end;

procedure TForm6.ButtonBuscarFDBClick(Sender: TObject);
begin
  OpenDialog.Filter := 'Banco Firebird|*.fdb';
  if OpenDialog.Execute then
    EditCaminhoFDB.Text := OpenDialog.FileName;
end;

procedure TForm6.ButtonImportarMarcasClick(Sender: TObject);
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfFile: TDbf;
  marcasUnicas: TStringList;
  descMarca, codMarcaStr: string;
  total, atual, novoCount: Integer;
  lastCodMarca, codMarca: Integer;
  usarCodMarca: Boolean;
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

  usarCodMarca := CheckUsarCodMarca.Checked and (ComboCodMarca.Text <> '');

  LabelProgresso.Caption := 'Importando marcas...';
  ProgressBar.Position := 0;
  Application.ProcessMessages;

  marcasUnicas := TStringList.Create;
  marcasUnicas.Sorted := True;
  marcasUnicas.Duplicates := dupIgnore;

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


      if not TableExists(query, 'MARCA') then
      begin
        ShowMessage('Tabela MARCA não encontrada no banco de dados. Criando tabela...');


        if trans.Active then
          trans.Commit;

        query.SQL.Text := 'CREATE TABLE MARCA (CODMARCA INTEGER PRIMARY KEY, DESCRICAO VARCHAR(50))';
        query.ExecSQL;
        trans.Commit;
      end;


      query.SQL.Text := 'SELECT COALESCE(MAX(CODMARCA), 0) AS N FROM MARCA';
      query.Open;
      lastCodMarca := StrToIntDef(Trim(query.Fields[0].AsString), 0);
      query.Close;


      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D FROM MARCA';
      query.Open;
      while not query.EOF do
      begin
        marcasUnicas.Add(query.FieldByName('D').AsString);
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
        descMarca := Trim(dbfFile.FieldByName(ComboDescricao.Text).AsString);


        if usarCodMarca then
        begin
          codMarcaStr := Trim(dbfFile.FieldByName(ComboCodMarca.Text).AsString);
          codMarca := StrToIntDef(codMarcaStr, 0);
        end
        else
        begin
          codMarca := 0;
        end;

        if (descMarca <> '') and (marcasUnicas.IndexOf(UpperCase(descMarca)) = -1) then
        begin

          if usarCodMarca and (codMarca > 0) then
          begin

            query.Close;
            query.SQL.Text := 'SELECT 1 FROM MARCA WHERE CODMARCA = :COD';
            query.ParamByName('COD').AsInteger := codMarca;
            query.Open;

            if not query.IsEmpty then
            begin

              LogMsg('log_erros_marcas.txt',
                Format('Código de marca já existe: %d - %s. Pulando registro.', [codMarca, descMarca]));
              query.Close;
              dbfFile.Next;
              Continue;
            end;
            query.Close;


            lastCodMarca := codMarca;
          end
          else
          begin

            Inc(lastCodMarca);
          end;

          query.SQL.Text := 'INSERT INTO MARCA (CODMARCA, DESCRICAO) VALUES (:COD, :DESC)';
          query.ParamByName('COD').AsInteger := lastCodMarca;
          query.ParamByName('DESC').AsString := Copy(descMarca, 1, 50);

          try
            query.ExecSQL;
            marcasUnicas.Add(UpperCase(descMarca));
            Inc(novoCount);

            if usarCodMarca and (codMarca > 0) then
              LogMsg('log_marcas_importadas.txt',
                Format('Marca importada com código específico: %d - %s', [codMarca, descMarca]))
            else
              LogMsg('log_marcas_importadas.txt',
                Format('Marca importada com código automático: %d - %s', [lastCodMarca, descMarca]));

          except
            on E: Exception do
            begin
              LogMsg('log_erros_marcas.txt',
                Format('Erro ao inserir marca: %s - %s', [descMarca, E.Message]));
            end;
          end;
        end;


        if (atual mod 100) = 0 then
        begin
          ProgressBar.Position := Round((atual / total) * 100);
          if usarCodMarca then
            LabelProgresso.Caption := Format('Processando: %d/%d - Novas: %d (com código)', [atual, total, novoCount])
          else
            LabelProgresso.Caption := Format('Processando: %d/%d - Novas: %d', [atual, total, novoCount]);
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

      if usarCodMarca then
        LabelProgresso.Caption := Format('✅ Marcas importadas: %d novas de %d registros (com código específico)', [novoCount, total])
      else
        LabelProgresso.Caption := Format('✅ Marcas importadas: %d novas de %d registros', [novoCount, total]);


      if novoCount > 0 then
      begin
        if usarCodMarca then
          ShowMessage(Format('Importação concluída!%sForam importadas %d novas marcas com códigos específicos.', [sLineBreak, novoCount]))
        else
          ShowMessage(Format('Importação concluída!%sForam importadas %d novas marcas.', [sLineBreak, novoCount]));
      end
      else
        ShowMessage('Nenhuma nova marca foi importada. Todas as marcas já existiam no banco.');

    except
      on E: Exception do
      begin
        if trans.Active then
        begin
          trans.Rollback;
        end;
        ShowMessage('Erro ao importar marcas: ' + E.Message);
        LogMsg('log_erros_marcas.txt', 'Erro geral importar marcas: ' + E.Message);
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

    if Assigned(marcasUnicas) then
      FreeAndNil(marcasUnicas);
  end;
end;

procedure TForm6.CarregarColunasDBF;
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
    ComboCodMarca.Clear;


    for i := 0 to dbfFile.FieldDefs.Count - 1 do
    begin
      ComboDescricao.Items.Add(dbfFile.FieldDefs[i].Name);
      ComboCodMarca.Items.Add(dbfFile.FieldDefs[i].Name);
    end;

    dbfFile.Close;


    if ComboDescricao.Items.Count > 0 then
      ComboDescricao.ItemIndex := 0;

    if ComboCodMarca.Items.Count > 0 then
      ComboCodMarca.ItemIndex := 0;

    ShowMessage('Colunas carregadas com sucesso!');
  finally
    dbfFile.Free;
  end;
end;

procedure TForm6.LogMsg(const nomeArq, msg: string);
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

function TForm6.SanitizeText(const S: string): string;
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

function TForm6.PrepareStringForDB(const S: string): string;
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

function TForm6.TableExists(Q: TSQLQuery; const tableName: string): Boolean;
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
