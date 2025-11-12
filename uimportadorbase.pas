unit uImportadorBase;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, uAnalisadorVariantes;

type
  TProgressEvent = procedure(APosicaoAtual: Integer) of object;
  TProgressLabelEvent = procedure(const ATextoProgresso: string) of object;
  TLogErroEvent = procedure(const ALinha: Integer; const AMensagemErro: string; const ADescricao: string) of object;

  TImportadorBase = class
  private
    FSituTributaria: string;

    FConn: TIBConnection;
    FQInsBarras: TSQLQuery;
    FUsarCodigoAlternativo: Boolean;
    FCodigoAlternativoField: string;
    FEditICMS: string;

    FMapeamento: TStringList;
    FUsedBarcodes: TStringList;
    FGradeCache: TStringList;
    FVarCache: TStringList;
    FPaiCache: TStringList;
    FGrupoCache: TStringList;
    FMarcaCache: TStringList;
    FICMSConfigCache: TStringList;

    FLastCodProduto: Integer;
    FLastIDGrade: Integer;
    FLastIDVariacao: Integer;
    FLastCodGrupo: Integer;
    FLastCodMarca: Integer;
    FLastID_ICMS: Integer;
    FPortaFirebird: Integer;

    FCaminhoDBF: string;
    FCaminhoFDB: string;
    FImportarEstoque: Boolean;
    FUsarGrade: Boolean;
    FLinhaInicial: Integer;
    FLinhasComErro: TStringList;

    FOnProgress: TProgressEvent;
    FOnProgressLabel: TProgressLabelEvent;
    FOnLogErro: TLogErroEvent;

    FQInsPai: TSQLQuery;
    FQInsSimples: TSQLQuery;
    FQInsFilho: TSQLQuery;
    FQInsGrade: TSQLQuery;
    FQInsVar: TSQLQuery;
    FQInsGrupo: TSQLQuery;
    FQInsMarca: TSQLQuery;
    FQInsICMS: TSQLQuery;

    procedure LogMsg(const nomeArq, msg: string);
    procedure LogErro(const ALinha: Integer; const AMensagemErro: string; const ADescricao: string);
    function RemoveAcentos(const S: string): string;
    function SanitizeText(const S: string): string;
    function PrepareStringForDB(const S: string): string;
    function Truncar(const texto: string; tamanho: Integer; const campo: string): string;
    procedure Prep(Q: TSQLQuery; const SQLText: string);
    function ExtractCODCFI(const descricao: string): string;
    function BuscarICMSNoBanco(const codCFI: string): Integer;
    function CodigoBarrasExiste(const codBarras: string; codProduto: Integer): Boolean;

    function TableExists(Q: TSQLQuery; const tableName: string): Boolean;
    procedure EnsureGradeTables(Q: TSQLQuery);
    procedure WarmUpCounters(Q: TSQLQuery);
    function NextCodigoProduto: Integer;
    function NextIDGrade: Integer;
    function NextIDVariacao: Integer;
    function NextCodigoGrupo: Integer;
    function NextCodigoMarca: Integer;
    function NextID_ICMS: Integer;
    function EnsureUniqueBarcode(Q: TSQLQuery; const candidate: string): string;

    function GetOrCreateGrade(trans: TSQLTransaction; const tipoGrade: string): Integer;
    function GetOrCreateGrupo(trans: TSQLTransaction; const descGrupo: string): Integer;
    function GetOrCreateMarca(trans: TSQLTransaction; const descMarca: string): Integer;
    function BuscarGrupoNoBanco(const descGrupo: string): Integer;
    function BuscarMarcaNoBanco(const descMarca: string): Integer;

    procedure PrepareAllQueries(conn: TIBConnection; trans: TSQLTransaction);
    procedure InserirCodigoAlternativo(trans: TSQLTransaction; codProduto: Integer; const codBarrasAlt: string);

    procedure ImportarProdutoNormal(trans: TSQLTransaction; const descricao, codBarras, un, ncm: string;
      precoVenda, precoCusto, precoEntrada, qtdEstoque: Double; codGrupo, codMarca, idICMS: Integer;
      const codBarrasAlt: string = '');

    procedure ImportarProdutoGrade(trans: TSQLTransaction; const descricao, codBarras, un, ncm: string;
      precoVenda, precoCusto, precoEntrada, qtdEstoque: Double; codGrupo, codMarca, idICMS: Integer;
      const codBarrasAlt: string = '');

    function GetOrCreateICMS(trans: TSQLTransaction; const configICMS: string): Integer;

    procedure DoProgress(Position: Integer);
    procedure DoProgressLabel(const Text: string);
    function CodigoBarrasIgualPrincipal(codProduto: Integer; const codBarrasAlt: string): Boolean;

  public
    property UsarCodigoAlternativo: Boolean read FUsarCodigoAlternativo write FUsarCodigoAlternativo;
    property CodigoAlternativoField: string read FCodigoAlternativoField write FCodigoAlternativoField;
    property PortaFirebird: Integer read FPortaFirebird write FPortaFirebird;
    property SituTributaria: string read FSituTributaria write FSituTributaria;
    property EditICMS: string read FEditICMS write FEditICMS;

    constructor Create;
    destructor Destroy; override;

    procedure ImportarProdutos;
    procedure ImportarSomenteGrupos;
    procedure ImportarSomenteMarcas;
    procedure ImportarSomenteTributacao;

    property Mapeamento: TStringList read FMapeamento;
    property CaminhoDBF: string read FCaminhoDBF write FCaminhoDBF;
    property CaminhoFDB: string read FCaminhoFDB write FCaminhoFDB;
    property ImportarEstoque: Boolean read FImportarEstoque write FImportarEstoque;
    property UsarGrade: Boolean read FUsarGrade write FUsarGrade;
    property LinhaInicial: Integer read FLinhaInicial write FLinhaInicial;
    property LinhasComErro: TStringList read FLinhasComErro;
    property OnProgress: TProgressEvent read FOnProgress write FOnProgress;
    property OnProgressLabel: TProgressLabelEvent read FOnProgressLabel write FOnProgressLabel;
    property OnLogErro: TLogErroEvent read FOnLogErro write FOnLogErro;
  end;

const
  BATCH_SIZE = 5000;
  UI_UPDATE_EVERY = 500;
  LOGS_ENABLED = True;
  CHUNK_SIZE = 1000;
  TEMP_SUBFOLDER = 'temp_dbf_chunks';

implementation

uses
  StrUtils, DateUtils, Variants, RegExpr;

{ TImportadorBase }

constructor TImportadorBase.Create;
begin
  inherited Create;
  FMapeamento := TStringList.Create;

  FUsedBarcodes := TStringList.Create;
  FUsedBarcodes.Sorted := True;
  FUsedBarcodes.Duplicates := dupIgnore;

  FGradeCache := TStringList.Create;
  FGradeCache.Sorted := True;
  FGradeCache.Duplicates := dupIgnore;

  FVarCache := TStringList.Create;
  FVarCache.Sorted := True;
  FVarCache.Duplicates := dupIgnore;

  FPaiCache := TStringList.Create;
  FPaiCache.Sorted := True;
  FPaiCache.Duplicates := dupIgnore;

  FGrupoCache := TStringList.Create;
  FGrupoCache.Sorted := True;
  FGrupoCache.Duplicates := dupIgnore;

  FMarcaCache := TStringList.Create;
  FMarcaCache.Sorted := True;
  FMarcaCache.Duplicates := dupIgnore;

  FICMSConfigCache := TStringList.Create;
  FICMSConfigCache.Sorted := True;
  FICMSConfigCache.Duplicates := dupIgnore;

  FLinhasComErro := TStringList.Create;
  FUsarGrade := False;
  FLinhaInicial := 1;
end;

destructor TImportadorBase.Destroy;
begin
  FQInsBarras.Free;
  FMapeamento.Free;
  FUsedBarcodes.Free;
  FGradeCache.Free;
  FVarCache.Free;
  FPaiCache.Free;
  FGrupoCache.Free;
  FMarcaCache.Free;
  FICMSConfigCache.Free;
  FLinhasComErro.Free;

  FQInsPai.Free;
  FQInsSimples.Free;
  FQInsFilho.Free;
  FQInsGrade.Free;
  FQInsVar.Free;
  FQInsGrupo.Free;
  FQInsMarca.Free;
  FQInsICMS.Free;

  inherited Destroy;
end;

procedure TImportadorBase.DoProgress(Position: Integer);
begin
  if Assigned(FOnProgress) then
    FOnProgress(Position);
end;

procedure TImportadorBase.DoProgressLabel(const Text: string);
begin
  if Assigned(FOnProgressLabel) then
    FOnProgressLabel(Text);
end;

procedure TImportadorBase.LogMsg(const nomeArq, msg: string);
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

procedure TImportadorBase.LogErro(const ALinha: Integer; const AMensagemErro: string; const ADescricao: string);
begin
  if Assigned(FOnLogErro) then
    FOnLogErro(ALinha, AMensagemErro, ADescricao);

  FLinhasComErro.Add(Format('Linha %d: %s - %s', [ALinha, AMensagemErro, ADescricao]));
  LogMsg('log_erros_detalhados.txt', Format('Linha %d: %s - %s', [ALinha, AMensagemErro, ADescricao]));
end;

function TImportadorBase.RemoveAcentos(const S: string): string;
const
  ComAcento = 'ÀÁÂãÄÅÇÈÉÊËÌÍÎÏÒÓÔÕÖÙÚÛÜÝàáâãäåçèéêëìíîïòóôõöùúûüýÿ';
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

function TImportadorBase.SanitizeText(const S: string): string;
var
  i: Integer;
  tmp: string;
begin
  tmp := UpperCase(RemoveAcentos(Trim(S)));

  Result := '';
  for i := 1 to Length(tmp) do
  begin
    if CharInSet(tmp[i], ['A'..'Z', '0'..'9', ' ', '-', '_', '/', '.', ',', '(', ')', ':']) then
      Result := Result + tmp[i];
  end;

  Result := Trim(StringReplace(Result, '  ', ' ', [rfReplaceAll]));
end;

function TImportadorBase.PrepareStringForDB(const S: string): string;
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

function TImportadorBase.Truncar(const texto: string; tamanho: Integer; const campo: string): string;
begin
  Result := Trim(texto);
  if Length(Result) > tamanho then
  begin
    Result := Copy(Result, 1, tamanho);
    LogMsg('log_truncados.txt', Format('Campo %s truncado: "%s"', [campo, texto]));
  end;
end;

procedure TImportadorBase.Prep(Q: TSQLQuery; const SQLText: string);
begin
  Q.Close;
  if Q.Prepared then
    Q.UnPrepare;
  Q.SQL.Text := SQLText;
  Q.Prepare;
end;

function TImportadorBase.TableExists(Q: TSQLQuery; const tableName: string): Boolean;
begin
  Prep(Q, 'SELECT 1 FROM RDB$RELATIONS WHERE UPPER(RDB$RELATION_NAME) = :T');
  Q.ParamByName('T').AsString := UpperCase(tableName);
  Q.Open;
  Result := not Q.IsEmpty;
  Q.Close;
end;

procedure TImportadorBase.EnsureGradeTables(Q: TSQLQuery);
begin
  if not TableExists(Q, 'TIPO_GRADE') then
  begin
    Q.SQL.Text :=
      'CREATE TABLE TIPO_GRADE (' +
      'CODGRADE INTEGER PRIMARY KEY, ' +
      'DESCRICAO VARCHAR(50), ' +
      'ATIVO CHAR(1))';
    Q.ExecSQL;
  end;

  if not TableExists(Q, 'TIPO_GRADE_VARIACOES') then
  begin
    Q.SQL.Text :=
      'CREATE TABLE TIPO_GRADE_VARIACOES (' +
      'CODGRADE INTEGER, ' +
      'IDVARIACAO INTEGER, ' +
      'DESCRICAO VARCHAR(20), ' +
      'PRIMARY KEY (CODGRADE, IDVARIACAO))';
    Q.ExecSQL;
  end;
end;

procedure TImportadorBase.WarmUpCounters(Q: TSQLQuery);
begin

  try
    Prep(Q, 'SELECT COALESCE(MAX(CODPRODUTO), 0) AS N FROM PRODUTO');
    Q.Open;
    if not Q.IsEmpty and not Q.Fields[0].IsNull then
      FLastCodProduto := StrToIntDef(Trim(Q.Fields[0].AsString), 0)
    else
      FLastCodProduto := 0;
    Q.Close;
  except
    on E: Exception do
    begin
      FLastCodProduto := 0;
      LogMsg('log_erros.txt', 'Erro ao buscar último produto: ' + E.Message);
    end;
  end;


  try
    Prep(Q, 'SELECT COALESCE(MAX(CODGRUPO), 0) AS N FROM GRUPO');
    Q.Open;
    if not Q.IsEmpty and not Q.Fields[0].IsNull then
    begin
      try
        FLastCodGrupo := StrToIntDef(Trim(Q.Fields[0].AsString), 0);
      except
        on E: Exception do
        begin
          FLastCodGrupo := 0;
          LogMsg('log_erros.txt', 'Erro converter CODGRUPO: ' + E.Message + ' - Valor: ' + Q.Fields[0].AsString);
        end;
      end;
    end
    else
      FLastCodGrupo := 0;
    Q.Close;
  except
    on E: Exception do
    begin
      FLastCodGrupo := 0;
      LogMsg('log_erros.txt', 'Erro ao buscar último grupo: ' + E.Message);
      try
        if not TableExists(Q, 'GRUPO') then
        begin
          Q.SQL.Text := 'CREATE TABLE GRUPO (CODGRUPO INTEGER PRIMARY KEY, DESCRICAO VARCHAR(50))';
          Q.ExecSQL;
          LogMsg('log_info.txt', 'Tabela GRUPO criada automaticamente');
        end;
      except
        on E2: Exception do
          LogMsg('log_erros.txt', 'Erro ao criar tabela GRUPO: ' + E2.Message);
      end;
    end;
  end;


  try
    Prep(Q, 'SELECT COALESCE(MAX(CODMARCA), 0) AS N FROM MARCA');
    Q.Open;
    if not Q.IsEmpty and not Q.Fields[0].IsNull then
    begin
      try
        FLastCodMarca := StrToIntDef(Trim(Q.Fields[0].AsString), 0);
      except
        on E: Exception do
        begin
          FLastCodMarca := 0;
          LogMsg('log_erros.txt', 'Erro converter CODMARCA: ' + E.Message + ' - Valor: ' + Q.Fields[0].AsString);
        end;
      end;
    end
    else
      FLastCodMarca := 0;
    Q.Close;
  except
    on E: Exception do
    begin
      FLastCodMarca := 0;
      LogMsg('log_erros.txt', 'Erro ao buscar última marca: ' + E.Message);
      try
        if not TableExists(Q, 'MARCA') then
        begin
          Q.SQL.Text := 'CREATE TABLE MARCA (CODMARCA INTEGER PRIMARY KEY, DESCRICAO VARCHAR(50))';
          Q.ExecSQL;
          LogMsg('log_info.txt', 'Tabela MARCA criada automaticamente');
        end;
      except
        on E2: Exception do
          LogMsg('log_erros.txt', 'Erro ao criar tabela MARCA: ' + E2.Message);
      end;
    end;
  end;


  try
    Prep(Q, 'SELECT COALESCE(MAX(ID_ICMS), 0) AS N FROM ICMS');
    Q.Open;
    if not Q.IsEmpty and not Q.Fields[0].IsNull then
    begin
      try
        FLastID_ICMS := StrToIntDef(Trim(Q.Fields[0].AsString), 0);
      except
        on E: Exception do
        begin
          FLastID_ICMS := 0;
          LogMsg('log_erros.txt', 'Erro converter ID_ICMS: ' + E.Message + ' - Valor: ' + Q.Fields[0].AsString);
        end;
      end;
    end
    else
      FLastID_ICMS := 0;
    Q.Close;
  except
    on E: Exception do
    begin
      FLastID_ICMS := 0;
      LogMsg('log_erros.txt', 'Erro ao buscar último ICMS: ' + E.Message);
      try
        if not TableExists(Q, 'ICMS') then
        begin
          Q.SQL.Text := 'CREATE TABLE ICMS (ID_ICMS INTEGER PRIMARY KEY, DESCRICAO VARCHAR(50), CST VARCHAR(3), ALIQICMS NUMERIC(15,2))';
          Q.ExecSQL;
          LogMsg('log_info.txt', 'Tabela ICMS criada automaticamente');
        end;
      except
        on E2: Exception do
          LogMsg('log_erros.txt', 'Erro ao criar tabela ICMS: ' + E2.Message);
      end;
    end;
  end;


  if TableExists(Q, 'TIPO_GRADE') then
  begin
    try
      Prep(Q, 'SELECT COALESCE(MAX(CODGRADE), 0) AS N FROM TIPO_GRADE');
      Q.Open;
      if not Q.IsEmpty and not Q.Fields[0].IsNull then
        FLastIDGrade := StrToIntDef(Trim(Q.Fields[0].AsString), 0)
      else
        FLastIDGrade := 0;
      Q.Close;
    except
      on E: Exception do
      begin
        FLastIDGrade := 0;
        LogMsg('log_erros.txt', 'Erro ao buscar última grade: ' + E.Message);
      end;
    end;
  end
  else
    FLastIDGrade := 0;


  if TableExists(Q, 'TIPO_GRADE_VARIACOES') then
  begin
    try
      Prep(Q, 'SELECT COALESCE(MAX(IDVARIACAO), 0) AS N FROM TIPO_GRADE_VARIACOES');
      Q.Open;
      if not Q.IsEmpty and not Q.Fields[0].IsNull then
        FLastIDVariacao := StrToIntDef(Trim(Q.Fields[0].AsString), 0)
      else
        FLastIDVariacao := 0;
      Q.Close;
    except
      on E: Exception do
      begin
        FLastIDVariacao := 0;
        LogMsg('log_erros.txt', 'Erro ao buscar última variação: ' + E.Message);
      end;
    end;
  end
  else
    FLastIDVariacao := 0;

  LogMsg('log_info.txt', Format(
    'Contadores inicializados - Produtos: %d, Grupos: %d, Marcas: %d, ICMS: %d, Grades: %d, Variações: %d',
    [FLastCodProduto, FLastCodGrupo, FLastCodMarca, FLastID_ICMS, FLastIDGrade, FLastIDVariacao]
  ));
end;

function TImportadorBase.NextCodigoProduto: Integer;
begin
  Inc(FLastCodProduto);
  Result := FLastCodProduto;
end;

function TImportadorBase.NextIDGrade: Integer;
begin
  Inc(FLastIDGrade);
  Result := FLastIDGrade;
end;

function TImportadorBase.NextIDVariacao: Integer;
begin
  Inc(FLastIDVariacao);
  Result := FLastIDVariacao;
end;

function TImportadorBase.NextCodigoGrupo: Integer;
begin
  Inc(FLastCodGrupo);
  Result := FLastCodGrupo;
end;

function TImportadorBase.NextCodigoMarca: Integer;
begin
  Inc(FLastCodMarca);
  Result := FLastCodMarca;
end;

function TImportadorBase.NextID_ICMS: Integer;
begin
  Inc(FLastID_ICMS);
  Result := FLastID_ICMS;
end;

function TImportadorBase.EnsureUniqueBarcode(Q: TSQLQuery; const candidate: string): string;
var
  c: string;
  i: Integer;
begin
  if Trim(candidate) <> '' then
    c := Copy(candidate, 1, 20)
  else
    c := 'SEM' + IntToStr(Random(9999999));

  for i := 0 to 9999 do
  begin
    if FUsedBarcodes.IndexOf(c) = -1 then
    begin
      FUsedBarcodes.Add(c);
      Exit(c);
    end;
    c := 'SEM' + IntToStr(Random(9999999));
  end;

  Result := 'SEM' + IntToStr(Random(999999999));
end;

function TImportadorBase.GetOrCreateGrade(trans: TSQLTransaction; const tipoGrade: string): Integer;
var
  key, s: string;
begin
  key := UpperCase(tipoGrade);
  s := FGradeCache.Values[key];
  if s <> '' then
    Exit(StrToIntDef(Trim(s), 0));

  Result := NextIDGrade;

  if (trans <> nil) and (not trans.Active) then
    trans.StartTransaction;

  with FQInsGrade.Params do
  begin
    ParamByName('C').AsInteger := Result;
    ParamByName('DS').AsString := Truncar(tipoGrade, 50, 'TIPO_GRADE.DESCRICAO');
  end;

  try
    FQInsGrade.ExecSQL;
    FGradeCache.Values[key] := IntToStr(Result);
  except
    on E: Exception do
    begin
      LogMsg('log_erros.txt', 'Erro criar grade: ' + E.Message + ' (GRADE=' + key + ')');
      if FGradeCache.Values[key] <> '' then
        Exit(StrToIntDef(Trim(FGradeCache.Values[key]), 0))
      else
        raise;
    end;
  end;
end;

function TImportadorBase.GetOrCreateGrupo(trans: TSQLTransaction; const descGrupo: string): Integer;
var
  key, s: string;
begin
  if descGrupo = '' then
    Exit(0);

  key := UpperCase(descGrupo);
  s := FGrupoCache.Values[key];
  if s <> '' then
    Exit(StrToIntDef(Trim(s), 0));

  Result := BuscarGrupoNoBanco(descGrupo);
  if Result > 0 then
  begin
    FGrupoCache.Values[key] := IntToStr(Result);
    Exit(Result);
  end;

  Result := 0;

  if descGrupo <> '' then
    LogMsg('log_info.txt', 'Grupo não encontrado: ' + descGrupo);
end;

function TImportadorBase.BuscarGrupoNoBanco(const descGrupo: string): Integer;
var
  qBusca: TSQLQuery;
  conn: TIBConnection;
  trans: TSQLTransaction;
begin
  Result := 0;

  if Assigned(FQInsGrupo) and Assigned(FQInsGrupo.DataBase) then
  begin
    qBusca := TSQLQuery.Create(nil);
    try
      qBusca.DataBase := FQInsGrupo.DataBase;
      qBusca.Transaction := FQInsGrupo.Transaction;

      with qBusca do
      begin
        SQL.Text := 'SELECT CODGRUPO FROM GRUPO WHERE UPPER(DESCRICAO) = UPPER(:DS)';
        Close;
        ParamByName('DS').AsString := descGrupo;
        Open;

        if not EOF then
          Result := FieldByName('CODGRUPO').AsInteger;

        Close;
      end;
    finally
      qBusca.Free;
    end;
  end
  else
  begin
    conn := TIBConnection.Create(nil);
    trans := TSQLTransaction.Create(nil);
    qBusca := TSQLQuery.Create(nil);

    try
      conn.DatabaseName := FCaminhoFDB;
      conn.UserName := 'sysdba';
      conn.Password := 'masterkey';
      conn.CharSet := 'UTF8';
      conn.Params.Add('lc_ctype=UTF8');
      conn.Params.Add('sql_dialect=3');
      conn.Transaction := trans;

      trans.DataBase := conn;
      qBusca.DataBase := conn;
      qBusca.Transaction := trans;

      conn.Open;

      with qBusca do
      begin
        SQL.Text := 'SELECT CODGRUPO FROM GRUPO WHERE UPPER(DESCRICAO) = UPPER(:DS)';
        Close;
        ParamByName('DS').AsString := descGrupo;
        Open;

        if not EOF then
          Result := FieldByName('CODGRUPO').AsInteger;

        Close;
      end;
    finally
      qBusca.Free;
      trans.Free;
      conn.Free;
    end;
  end;
end;

function TImportadorBase.GetOrCreateMarca(trans: TSQLTransaction; const descMarca: string): Integer;
var
  key, s: string;
begin
  if descMarca = '' then
    Exit(0);

  key := UpperCase(descMarca);
  s := FMarcaCache.Values[key];
  if s <> '' then
    Exit(StrToIntDef(Trim(s), 0));

  Result := BuscarMarcaNoBanco(descMarca);
  if Result > 0 then
  begin
    FMarcaCache.Values[key] := IntToStr(Result);
    Exit(Result);
  end;

  Result := 0;

  if descMarca <> '' then
    LogMsg('log_info.txt', 'Marca não encontrada: ' + descMarca);
end;

function TImportadorBase.BuscarMarcaNoBanco(const descMarca: string): Integer;
var
  qBusca: TSQLQuery;
  conn: TIBConnection;
  trans: TSQLTransaction;
begin
  Result := 0;

  if Trim(descMarca) = '' then
    Exit;

  if Assigned(FQInsMarca) and Assigned(FQInsMarca.DataBase) and FQInsMarca.DataBase.Connected then
  begin
    qBusca := TSQLQuery.Create(nil);
    try
      qBusca.DataBase := FQInsMarca.DataBase;
      qBusca.Transaction := FQInsMarca.Transaction;

      with qBusca do
      begin
        SQL.Text := 'SELECT CODMARCA FROM MARCA WHERE UPPER(TRIM(DESCRICAO)) = UPPER(TRIM(:DS))';
        Close;
        ParamByName('DS').AsString := descMarca;
        Open;

        if not EOF then
          Result := FieldByName('CODMARCA').AsInteger;

        Close;
      end;
    except
      on E: Exception do
      begin
        LogMsg('log_erros.txt', 'Erro ao buscar marca (conexão existente): ' + E.Message);
        Result := 0;
      end;
    end;
    qBusca.Free;
  end
  else
  begin
    conn := TIBConnection.Create(nil);
    trans := TSQLTransaction.Create(nil);
    qBusca := TSQLQuery.Create(nil);

    try
      try
        conn.DatabaseName := FCaminhoFDB;
        conn.UserName := 'sysdba';
        conn.Password := 'masterkey';
        conn.CharSet := 'UTF8';
        conn.Params.Add('lc_ctype=UTF8');
        conn.Params.Add('sql_dialect=3');
        conn.Transaction := trans;

        trans.DataBase := conn;
        qBusca.DataBase := conn;
        qBusca.Transaction := trans;

        conn.Open;

        with qBusca do
        begin
          SQL.Text := 'SELECT CODMARCA FROM MARCA WHERE UPPER(TRIM(DESCRICAO)) = UPPER(TRIM(:DS))';
          Close;
          ParamByName('DS').AsString := descMarca;
          Open;

          if not EOF then
            Result := FieldByName('CODMARCA').AsInteger;

          Close;
        end;

      except
        on E: Exception do
        begin
          LogMsg('log_erros.txt', 'Erro ao buscar marca (conexão temp): ' + E.Message);
          Result := 0;
        end;
      end;
    finally
      if qBusca.Active then qBusca.Close;
      qBusca.Free;

      if trans.Active then trans.Rollback;
      trans.Free;

      if conn.Connected then conn.Close;
      conn.Free;
    end;
  end;
end;

function TImportadorBase.GetOrCreateICMS(trans: TSQLTransaction; const configICMS: string): Integer;
var
  key, s: string;
begin
  if configICMS = '' then
    Exit(0);

  key := UpperCase(configICMS);
  s := FICMSConfigCache.Values[key];
  if s <> '' then
    Exit(StrToIntDef(Trim(s), 0));

  Result := NextID_ICMS;

  if (trans <> nil) and (not trans.Active) then
    trans.StartTransaction;

  with FQInsICMS.Params do
  begin
    ParamByName('ID').AsInteger := Result;
    ParamByName('DS').AsString := Truncar('ICMS ' + configICMS, 50, 'ICMS.DESCRICAO');
    ParamByName('CST').AsString := '00';
    ParamByName('ALIQ').AsFloat := 17.0;
    ParamByName('ORIGEM').AsInteger := 0;
  end;

  try
    FQInsICMS.ExecSQL;
    FICMSConfigCache.Values[key] := IntToStr(Result);
  except
    on E: Exception do
    begin
      LogMsg('log_erros.txt', 'Erro criar ICMS: ' + E.Message + ' (ICMS=' + key + ')');
      if FICMSConfigCache.Values[key] <> '' then
        Exit(StrToIntDef(Trim(FICMSConfigCache.Values[key]), 0))
      else
        raise;
    end;
  end;
end;

procedure TImportadorBase.ImportarProdutoNormal(trans: TSQLTransaction; const descricao, codBarras, un, ncm: string;
  precoVenda, precoCusto, precoEntrada, qtdEstoque: Double; codGrupo, codMarca, idICMS: Integer;
  const codBarrasAlt: string = '');
var
  codProduto: Integer;
begin
  codProduto := NextCodigoProduto;

  with FQInsSimples.Params do
  begin
    ParamByName('C').AsInteger := codProduto;
    ParamByName('D').AsString := PrepareStringForDB(descricao);
    ParamByName('B').AsString := PrepareStringForDB(codBarras);
    ParamByName('U').AsString := PrepareStringForDB(un);
    ParamByName('N').AsString := PrepareStringForDB(ncm);
    ParamByName('PV').AsFloat := precoVenda;
    ParamByName('PC').AsFloat := precoCusto;
    ParamByName('PE').AsFloat := precoEntrada;
    ParamByName('E').AsFloat := qtdEstoque;
    ParamByName('GRUPO').AsInteger := codGrupo;
    ParamByName('MARCA').AsInteger := codMarca;
    if Trim(FEditICMS) <> '' then
      ParamByName('ICMS').AsInteger := StrToIntDef(FEditICMS, idICMS)
    else
      ParamByName('ICMS').AsInteger := idICMS;
    ParamByName('SITU_TRIBUTA').AsString := FSituTributaria;
  end;

  FQInsSimples.ExecSQL;

  InserirCodigoAlternativo(trans, codProduto, codBarrasAlt);
end;

procedure TImportadorBase.ImportarProdutoGrade(trans: TSQLTransaction; const descricao, codBarras, un, ncm: string;
  precoVenda, precoCusto, precoEntrada, qtdEstoque: Double; codGrupo, codMarca, idICMS: Integer;
  const codBarrasAlt: string = '');
var
  analisador: TAnalisadorVariantes;
  baseProduto, variacao, tipoGrade: string;
  codGrade, idVar, codPai, codFilho: Integer;
  codBarrasChild: string;
  usandoIA: Boolean;
begin
  analisador := TAnalisadorVariantes.Create(descricao);
  try
    usandoIA := analisador.TemVariacao;

    if usandoIA then
    begin
      baseProduto := analisador.GetBaseProduto;
      variacao := analisador.GetVariacaoPrincipal;
      tipoGrade := analisador.GetTipoVariacao;
    end
    else
    begin
      ImportarProdutoNormal(trans, descricao, codBarras, un, ncm, precoVenda,
        precoCusto, precoEntrada, qtdEstoque, codGrupo, codMarca, idICMS, codBarrasAlt);
      Exit;
    end;

    codGrade := GetOrCreateGrade(trans, tipoGrade);

    if FPaiCache.Values[UpperCase(baseProduto)] <> '' then
      codPai := StrToIntDef(Trim(FPaiCache.Values[UpperCase(baseProduto)]), 0)
    else
    begin
      codPai := NextCodigoProduto;
      with FQInsPai.Params do
      begin
        ParamByName('C').AsInteger := codPai;
        ParamByName('D').AsString := PrepareStringForDB(baseProduto);
        ParamByName('G').AsInteger := codGrade;
        ParamByName('B').AsString := PrepareStringForDB(codBarras);
        ParamByName('U').AsString := PrepareStringForDB(un);
        ParamByName('N').AsString := PrepareStringForDB(ncm);
        ParamByName('PV').AsFloat := precoVenda;
        ParamByName('PC').AsFloat := precoCusto;
        ParamByName('PE').AsFloat := precoEntrada;
        ParamByName('GRUPO').AsInteger := codGrupo;
        ParamByName('MARCA').AsInteger := codMarca;
        ParamByName('ICMS').AsInteger := idICMS;
      end;
      FQInsPai.ExecSQL;
      FPaiCache.Values[UpperCase(baseProduto)] := IntToStr(codPai);

      // Inserir código alternativo para o produto pai
      InserirCodigoAlternativo(trans, codPai, codBarrasAlt);
    end;

    if FVarCache.Values[IntToStr(codGrade) + '|' + UpperCase(variacao)] <> '' then
      idVar := StrToIntDef(Trim(FVarCache.Values[IntToStr(codGrade) + '|' + UpperCase(variacao)]), 0)
    else
    begin
      idVar := NextIDVariacao;
      with FQInsVar.Params do
      begin
        ParamByName('G').AsInteger := codGrade;
        ParamByName('I').AsInteger := idVar;
        ParamByName('DS').AsString := PrepareStringForDB(Truncar(variacao, 20, 'TIPO_GRADE_VARIACOES.DESCRICAO'));
      end;
      FQInsVar.ExecSQL;
      FVarCache.Values[IntToStr(codGrade) + '|' + UpperCase(variacao)] := IntToStr(idVar);
    end;

    codFilho := NextCodigoProduto;
    codBarrasChild := EnsureUniqueBarcode(nil, codBarras + '_' + IntToStr(codPai));

    with FQInsFilho.Params do
    begin
      ParamByName('C').AsInteger := codFilho;
      ParamByName('D').AsString := PrepareStringForDB(descricao);
      ParamByName('GEN').AsInteger := codPai;
      ParamByName('G1').AsInteger := codGrade;
      ParamByName('G2').AsInteger := codGrade;
      ParamByName('V1').AsInteger := idVar;
      ParamByName('V2').AsInteger := idVar;
      ParamByName('B').AsString := PrepareStringForDB(codBarrasChild);
      ParamByName('U').AsString := PrepareStringForDB(un);
      ParamByName('N').AsString := PrepareStringForDB(ncm);
      ParamByName('PV').AsFloat := precoVenda;
      ParamByName('PC').AsFloat := precoCusto;
      ParamByName('PE').AsFloat := precoEntrada;
      ParamByName('GRUPO').AsInteger := codGrupo;
      ParamByName('MARCA').AsInteger := codMarca;
      ParamByName('ICMS').AsInteger := idICMS;
    end;
    FQInsFilho.ExecSQL;

    // Inserir código alternativo para o produto filho
    InserirCodigoAlternativo(trans, codFilho, codBarrasAlt);

  finally
    analisador.Free;
  end;
end;

function TImportadorBase.ExtractCODCFI(const descricao: string): string;
var
  parts: TStringList;
  i: Integer;
begin
  Result := '';
  parts := TStringList.Create;
  try
    parts.Delimiter := ' ';
    parts.DelimitedText := descricao;


    for i := 0 to parts.Count - 2 do
    begin
      if UpperCase(parts[i]) = 'CF' then
      begin
        Result := parts[i + 1];
        Break;
      end;
    end;
  finally
    parts.Free;
  end;
end;

function TImportadorBase.BuscarICMSNoBanco(const codCFI: string): Integer;
var
  qBusca: TSQLQuery;
begin
  Result := 0;

  if Assigned(FQInsICMS) and Assigned(FQInsICMS.DataBase) then
  begin
    qBusca := TSQLQuery.Create(nil);
    try
      qBusca.DataBase := FQInsICMS.DataBase;
      qBusca.Transaction := FQInsICMS.Transaction;

      qBusca.SQL.Text := 'SELECT ID_ICMS FROM ICMS WHERE UPPER(DESCRICAO) LIKE :CF';
      qBusca.ParamByName('CF').AsString := '%CF ' + codCFI + '%';
      qBusca.Open;

      if not qBusca.EOF then
        Result := qBusca.FieldByName('ID_ICMS').AsInteger;

      qBusca.Close;
    finally
      qBusca.Free;
    end;
  end;
end;

function TImportadorBase.CodigoBarrasExiste(const codBarras: string; codProduto: Integer): Boolean;
var
  qBusca: TSQLQuery;
begin
  Result := False;

  qBusca := TSQLQuery.Create(nil);
  try
    qBusca.DataBase := FConn;
    qBusca.SQL.Text := 'SELECT COUNT(*) FROM BARRAS WHERE CODBARRAS = :CODBARRAS AND CODPRODUTO <> :CODPRODUTO';
    qBusca.ParamByName('CODBARRAS').AsString := codBarras;
    qBusca.ParamByName('CODPRODUTO').AsInteger := codProduto;
    qBusca.Open;

    Result := qBusca.Fields[0].AsInteger > 0;
    qBusca.Close;
  finally
    qBusca.Free;
  end;
end;

function TImportadorBase.CodigoBarrasIgualPrincipal(codProduto: Integer; const codBarrasAlt: string): Boolean;
var
  qBusca: TSQLQuery;
  codBarrasPrincipal: string;
begin
  Result := False;

  if not Assigned(FConn) then
    Exit;

  qBusca := TSQLQuery.Create(nil);
  try
    qBusca.DataBase := FConn;
    qBusca.SQL.Text := 'SELECT CODBARRAS FROM PRODUTO WHERE CODPRODUTO = :CODPRODUTO';
    qBusca.ParamByName('CODPRODUTO').AsInteger := codProduto;
    qBusca.Open;

    if not qBusca.EOF then
    begin
      codBarrasPrincipal := Trim(qBusca.FieldByName('CODBARRAS').AsString);
      Result := (codBarrasPrincipal = codBarrasAlt);

      if Result then
        LogMsg('log_info.txt', 'Código alternativo igual ao principal: ' + codBarrasAlt + ' (Produto: ' + IntToStr(codProduto) + ')');
    end;

    qBusca.Close;
  finally
    qBusca.Free;
  end;
end;

procedure TImportadorBase.InserirCodigoAlternativo(trans: TSQLTransaction; codProduto: Integer; const codBarrasAlt: string);
var
  codBarrasClean: string;
begin
  // Verificação inicial
  if not FUsarCodigoAlternativo or (Trim(codBarrasAlt) = '') then
    Exit;

  codBarrasClean := PrepareStringForDB(SanitizeText(codBarrasAlt));

  if codBarrasClean = '' then
    Exit;

  // ⚠️ VERIFICAÇÃO NOVA: Se o código alternativo for igual ao código de barras principal, NÃO insere
  if CodigoBarrasIgualPrincipal(codProduto, codBarrasClean) then
  begin
    LogMsg('log_avisos.txt', 'Código alternativo igual ao código principal: ' + codBarrasClean +
      ' (Produto: ' + IntToStr(codProduto) + ') - Ignorando inserção em BARRAS');
    Exit;
  end;

  // Verificações de segurança
  if Length(codBarrasClean) < 3 then
  begin
    LogMsg('log_avisos.txt', 'Código alternativo muito curto: ' + codBarrasClean +
      ' (Produto: ' + IntToStr(codProduto) + ')');
    Exit;
  end;

  if UpperCase(Copy(codBarrasClean, 1, 3)) = 'SEM' then
  begin
    LogMsg('log_avisos.txt', 'Código alternativo inválido (SEM): ' + codBarrasClean +
      ' (Produto: ' + IntToStr(codProduto) + ')');
    Exit;
  end;

  // Verificar se o código de barras já existe ANTES de tentar inserir
  if CodigoBarrasExiste(codBarrasClean, codProduto) then
  begin
    LogMsg('log_avisos.txt', 'Código alternativo já existe: ' + codBarrasClean +
      ' (Produto: ' + IntToStr(codProduto) + ')');
    Exit;
  end;

  // Tentar inserir e capturar exceções
  with FQInsBarras.Params do
  begin
    ParamByName('CODBARRAS').AsString := Truncar(codBarrasClean, 20, 'BARRAS.CODBARRAS');
    ParamByName('CODPRODUTO').AsInteger := codProduto;
    ParamByName('DESC_ACRES').AsString := 'A';
    ParamByName('PORCENTAGEM').AsFloat := 0;
    ParamByName('QTDEMBALAGEM').AsFloat := 1;
    ParamByName('ALTERADO').AsDateTime := Now;
  end;

  try
    FQInsBarras.ExecSQL;
    // Log apenas para debug - pode remover depois
    LogMsg('log_info.txt', 'Código alternativo inserido com sucesso: ' + codBarrasClean +
      ' (Produto: ' + IntToStr(codProduto) + ')');
  except
    on E: Exception do
    begin
      // Se for erro de duplicata (incluindo da trigger), apenas logar como aviso
      if (Pos('BARRAS_COD_EXISTENTE', E.Message) > 0) or
         (Pos('BI_BARRAS_CODBARRAS', E.Message) > 0) or  // Nome correto da trigger
         (Pos('BARRAS_CODBARRAS', E.Message) > 0) or
         (Pos('duplicate', LowerCase(E.Message)) > 0) or
         (Pos('unique', LowerCase(E.Message)) > 0) or
         (Pos('já existe', LowerCase(E.Message)) > 0) then
      begin
        LogMsg('log_avisos.txt', 'Código alternativo já existe: ' + codBarrasClean +
          ' (Produto: ' + IntToStr(codProduto) + ') - ' + E.Message);
      end
      else
      begin
        // Outros erros são logados como erro
        LogMsg('log_erros.txt', 'Erro ao inserir código alternativo: ' + E.Message +
          ' (Produto: ' + IntToStr(codProduto) + ', Código: ' + codBarrasClean + ')');
        // Relançar a exceção se for um erro não esperado
        raise;
      end;
    end;
  end;
end;

procedure TImportadorBase.PrepareAllQueries(conn: TIBConnection; trans: TSQLTransaction);
begin
  FConn := conn; // Armazena a conexão para uso na função CodigoBarrasExiste

  FQInsSimples := TSQLQuery.Create(nil);
  FQInsSimples.DataBase := conn;
  FQInsSimples.Transaction := trans;

  FQInsSimples.SQL.Text :=
    'INSERT INTO PRODUTO (CODPRODUTO, DESCRICAO, CODBARRAS, UN, NCM, PRECO_VENDA, PRECO_CUSTO, PRECO_ENTRADA, ' +
    'QTD_ESTOQUE, GRADE, SITU_TRIBUTA, CODGRUPO, CODMARCA, ID_ICMS, FLAG) ' +
    'VALUES (:C, :D, :B, :U, :N, :PV, :PC, :PE, :E, ''N'', :SITU_TRIBUTA, :GRUPO, :MARCA, :ICMS, 0)';
  FQInsSimples.Prepare;

  FQInsPai := TSQLQuery.Create(nil);
  FQInsPai.DataBase := conn;
  FQInsPai.Transaction := trans;
  FQInsPai.SQL.Text :=
    'INSERT INTO PRODUTO (CODPRODUTO, DESCRICAO, GRADE, CODGRADE, CODBARRAS, UN, NCM, PRECO_VENDA, PRECO_CUSTO, ' +
    'PRECO_ENTRADA, SITU_TRIBUTA, CODGRUPO, CODMARCA, ID_ICMS, FLAG) ' +
    'VALUES (:C, :D, ''S'', :G, :B, :U, :N, :PV, :PC, :PE, :SITU_TRIBUTA, :GRUPO, :MARCA, :ICMS, 0)';
  FQInsPai.Prepare;

  FQInsFilho := TSQLQuery.Create(nil);
  FQInsFilho.DataBase := conn;
  FQInsFilho.Transaction := trans;
  FQInsFilho.SQL.Text :=
    'INSERT INTO PRODUTO (' +
    '  CODPRODUTO, DESCRICAO, GRADE, CODPRODUTO_GENERICO, ' +
    '  CODGRADE, CODGRADE2, IDVARIACAO, IDVARIACAO2, ' +
    '  CODBARRAS, UN, NCM, PRECO_VENDA, PRECO_CUSTO, PRECO_ENTRADA, SITU_TRIBUTA, ' +
    '  CODGRUPO, CODMARCA, ID_ICMS, FLAG ' +
    ') VALUES (' +
    '  :C, :D, ''F'', :GEN, ' +
    '  :G1, :G2, :V1, :V2, ' +
    '  :B, :U, :N, :PV, :PC, :PE, :SITU_TRIBUTA, ' +
    '  :GRUPO, :MARCA, :ICMS, 0)';
  FQInsFilho.Prepare;

  FQInsBarras := TSQLQuery.Create(nil);
  FQInsBarras.DataBase := conn;
  FQInsBarras.Transaction := trans;
  FQInsBarras.SQL.Text :=
    'INSERT INTO BARRAS (CODBARRAS, CODPRODUTO, DESC_ACRES, PORCENTAGEM, QTDEMBALAGEM, ALTERADO) ' +
    'VALUES (:CODBARRAS, :CODPRODUTO, :DESC_ACRES, :PORCENTAGEM, :QTDEMBALAGEM, :ALTERADO)';
  FQInsBarras.Prepare;

  FQInsGrade := TSQLQuery.Create(nil);
  FQInsGrade.DataBase := conn;
  FQInsGrade.Transaction := trans;
  FQInsGrade.SQL.Text := 'INSERT INTO TIPO_GRADE (CODGRADE, DESCRICAO, ATIVO) VALUES (:C, :DS, ''S'')';
  FQInsGrade.Prepare;

  FQInsVar := TSQLQuery.Create(nil);
  FQInsVar.DataBase := conn;
  FQInsVar.Transaction := trans;
  FQInsVar.SQL.Text := 'INSERT INTO TIPO_GRADE_VARIACOES (CODGRADE, IDVARIACAO, DESCRICAO) VALUES (:G, :I, :DS)';
  FQInsVar.Prepare;

  FQInsGrupo := TSQLQuery.Create(nil);
  FQInsGrupo.DataBase := conn;
  FQInsGrupo.Transaction := trans;
  FQInsGrupo.SQL.Text := 'INSERT INTO GRUPO (CODGRUPO, DESCRICAO) VALUES (:C, :DS)';
  FQInsGrupo.Prepare;

  FQInsMarca := TSQLQuery.Create(nil);
  FQInsMarca.DataBase := conn;
  FQInsMarca.Transaction := trans;
  FQInsMarca.SQL.Text := 'INSERT INTO MARCA (CODMARCA, DESCRICAO) VALUES (:C, :DS)';
  FQInsMarca.Prepare;

  FQInsICMS := TSQLQuery.Create(nil);
  FQInsICMS.DataBase := conn;
  FQInsICMS.Transaction := trans;
  FQInsICMS.SQL.Text :=
    'INSERT INTO ICMS (ID_ICMS, DESCRICAO, CST, ALIQICMS, ORIGEM) VALUES (:ID, :DS, :CST, :ALIQ, :ORIGEM)';
  FQInsICMS.Prepare;
end;

procedure TImportadorBase.ImportarProdutos;
var
  descRaw, desc, codBarrasRaw, codBarras, unRaw, un, ncmRaw, ncm: string;
  grupoRaw, grupo, marcaRaw, marca, codCFI: string;
  precoVenda, precoCusto, precoEntrada, qtdEstoque: Double;
  codGrupo, codMarca, idICMS: Integer;
  total, atual, numLinha: Integer;
  t0, t1: TDateTime;
  selectedFields: TStringList;
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfProdutos: TDbf;
  totalGeral: Integer;
  codBarrasAltRaw, codBarrasAlt: string;
begin
  Randomize;
  FLinhasComErro.Clear;

  trans := TSQLTransaction.Create(nil);
  query := TSQLQuery.Create(nil);
  dbfProdutos := nil;

  try
    DoProgressLabel('Conectando ao banco...');

    conn := TIBConnection.Create(nil);
    conn.DatabaseName := FCaminhoFDB;
    conn.UserName := 'sysdba';
    conn.Password := 'masterkey';
    conn.Port := FPortaFirebird;
    conn.CharSet := 'UTF8';
    conn.Params.Add('lc_ctype=UTF8');
    conn.Params.Add('sql_dialect=3');
    conn.Transaction := trans;

    trans.DataBase := conn;
    query.DataBase := conn;
    query.Transaction := trans;

    conn.Open;
    trans.StartTransaction;

    DoProgressLabel('Preparando ambiente...');

    if FUsarGrade then
      EnsureGradeTables(query);

    WarmUpCounters(query);
    PrepareAllQueries(conn, trans);

    FUsedBarcodes.Clear;
    query.SQL.Text := 'SELECT CODBARRAS FROM PRODUTO WHERE CODBARRAS IS NOT NULL';
    query.Open;
    while not query.EOF do
    begin
      if Trim(query.Fields[0].AsString) <> '' then
        FUsedBarcodes.Add(query.Fields[0].AsString);
      query.Next;
    end;
    query.Close;

    if FUsarGrade then
    begin
      FPaiCache.Clear;
      FPaiCache.Sorted := False;
      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D, CODPRODUTO FROM PRODUTO WHERE GRADE=''S''';
      query.Open;
      while not query.EOF do
      begin
        FPaiCache.Values[query.FieldByName('D').AsString] :=
          IntToStr(StrToIntDef(Trim(query.FieldByName('CODPRODUTO').AsString), 0));
        query.Next;
      end;
      query.Close;
      FPaiCache.Sorted := True;

      FGradeCache.Clear;
      FGradeCache.Sorted := False;
      query.SQL.Text := 'SELECT CODGRADE, UPPER(DESCRICAO) AS D FROM TIPO_GRADE';
      query.Open;
      while not query.EOF do
      begin
        FGradeCache.Values[query.FieldByName('D').AsString] :=
          IntToStr(StrToIntDef(Trim(query.FieldByName('CODGRADE').AsString), 0));
        query.Next;
      end;
      query.Close;
      FGradeCache.Sorted := True;

      FVarCache.Clear;
      FVarCache.Sorted := False;
      query.SQL.Text := 'SELECT CODGRADE, IDVARIACAO, UPPER(DESCRICAO) AS D FROM TIPO_GRADE_VARIACOES';
      query.Open;
      while not query.EOF do
      begin
        FVarCache.Values[
          IntToStr(StrToIntDef(Trim(query.FieldByName('CODGRADE').AsString), 0)) + '|' +
          query.FieldByName('D').AsString
        ] := IntToStr(StrToIntDef(Trim(query.FieldByName('IDVARIACAO').AsString), 0));
        query.Next;
      end;
      query.Close;
      FVarCache.Sorted := True;
    end;

    FGrupoCache.Clear;
    FGrupoCache.Sorted := False;
    query.SQL.Text := 'SELECT CODGRUPO, UPPER(DESCRICAO) AS D FROM GRUPO';
    query.Open;
    while not query.EOF do
    begin
      FGrupoCache.Values[query.FieldByName('D').AsString] :=
        IntToStr(StrToIntDef(Trim(query.FieldByName('CODGRUPO').AsString), 0));
      query.Next;
    end;
    query.Close;
    FGrupoCache.Sorted := True;

    FMarcaCache.Clear;
    FMarcaCache.Sorted := False;
    query.SQL.Text := 'SELECT CODMARCA, UPPER(DESCRICAO) AS D FROM MARCA';
    query.Open;
    while not query.EOF do
    begin
      FMarcaCache.Values[query.FieldByName('D').AsString] :=
        IntToStr(StrToIntDef(Trim(query.FieldByName('CODMARCA').AsString), 0));
      query.Next;
    end;
    query.Close;
    FMarcaCache.Sorted := True;

    FICMSConfigCache.Clear;
    FICMSConfigCache.Sorted := False;
    if TableExists(query, 'ICMS') then
    begin
      query.SQL.Text := 'SELECT ID_ICMS, DESCRICAO FROM ICMS';
      query.Open;
      while not query.EOF do
      begin
        FICMSConfigCache.Values[ExtractCODCFI(query.FieldByName('DESCRICAO').AsString)] :=
          IntToStr(StrToIntDef(Trim(query.FieldByName('ID_ICMS').AsString), 0));
        query.Next;
      end;
      query.Close;
    end;
    FICMSConfigCache.Sorted := True;

    selectedFields := TStringList.Create;
    try
      if FMapeamento.Values['DESCRICAO'] <> '' then selectedFields.Add(FMapeamento.Values['DESCRICAO']);
      if FMapeamento.Values['BARRAS'] <> '' then selectedFields.Add(FMapeamento.Values['BARRAS']);
      if FMapeamento.Values['UN'] <> '' then selectedFields.Add(FMapeamento.Values['UN']);
      if FMapeamento.Values['NCM'] <> '' then selectedFields.Add(FMapeamento.Values['NCM']);
      if FMapeamento.Values['PRECO_VENDA'] <> '' then selectedFields.Add(FMapeamento.Values['PRECO_VENDA']);
      if FMapeamento.Values['PRECO_CUSTO'] <> '' then selectedFields.Add(FMapeamento.Values['PRECO_CUSTO']);
      if FMapeamento.Values['PRECO_ENTRADA'] <> '' then selectedFields.Add(FMapeamento.Values['PRECO_ENTRADA']);
      if FImportarEstoque and (FMapeamento.Values['ESTOQUE'] <> '') then
        selectedFields.Add(FMapeamento.Values['ESTOQUE']);
      if FMapeamento.Values['GRUPO'] <> '' then selectedFields.Add(FMapeamento.Values['GRUPO']);
      if FMapeamento.Values['MARCA'] <> '' then selectedFields.Add(FMapeamento.Values['MARCA']);

      dbfProdutos := TDbf.Create(nil);
      try
        dbfProdutos.FilePathFull := ExtractFilePath(FCaminhoDBF);
        dbfProdutos.TableName := ExtractFileName(FCaminhoDBF);
        dbfProdutos.Open;

        totalGeral := dbfProdutos.RecordCount;
        atual := 0;
        numLinha := 0;
        t0 := Now;

        DoProgress(0);
        DoProgressLabel(Format('Processando %d registros...', [totalGeral]));

        dbfProdutos.First;
        while not dbfProdutos.EOF do
        begin
          Inc(numLinha);
          Inc(atual);

          if numLinha < FLinhaInicial then
          begin
            dbfProdutos.Next;
            Continue;
          end;

          if FUsarCodigoAlternativo and (FCodigoAlternativoField <> '') then
            codBarrasAltRaw := dbfProdutos.FieldByName(FCodigoAlternativoField).AsString
          else
            codBarrasAltRaw := '';

          codBarrasAlt := PrepareStringForDB(SanitizeText(codBarrasAltRaw));

          if (atual mod UI_UPDATE_EVERY) = 0 then
          begin
            if totalGeral > 0 then
              DoProgress(Round((atual / totalGeral) * 100));

            DoProgressLabel(Format('Importando %d/%d...', [atual, totalGeral]));
          end;

          try
            descRaw := dbfProdutos.FieldByName(FMapeamento.Values['DESCRICAO']).AsString;
            codBarrasRaw := dbfProdutos.FieldByName(FMapeamento.Values['BARRAS']).AsString;
            unRaw := dbfProdutos.FieldByName(FMapeamento.Values['UN']).AsString;

            precoVenda := dbfProdutos.FieldByName(FMapeamento.Values['PRECO_VENDA']).AsFloat;

            if FMapeamento.Values['NCM'] <> '' then
              ncmRaw := dbfProdutos.FieldByName(FMapeamento.Values['NCM']).AsString
            else
              ncmRaw := '';

            if FMapeamento.Values['PRECO_CUSTO'] <> '' then
              precoCusto := dbfProdutos.FieldByName(FMapeamento.Values['PRECO_CUSTO']).AsFloat
            else
              precoCusto := 0;

            if FMapeamento.Values['PRECO_ENTRADA'] <> '' then
              precoEntrada := dbfProdutos.FieldByName(FMapeamento.Values['PRECO_ENTRADA']).AsFloat
            else
              precoEntrada := 0;

            if FImportarEstoque and (FMapeamento.Values['ESTOQUE'] <> '') then
              qtdEstoque := dbfProdutos.FieldByName(FMapeamento.Values['ESTOQUE']).AsFloat
            else
              qtdEstoque := 0;

            if FMapeamento.Values['GRUPO'] <> '' then
              grupoRaw := dbfProdutos.FieldByName(FMapeamento.Values['GRUPO']).AsString
            else
              grupoRaw := '';

            if FMapeamento.Values['MARCA'] <> '' then
              marcaRaw := dbfProdutos.FieldByName(FMapeamento.Values['MARCA']).AsString
            else
              marcaRaw := '';

            codCFI := '';
            if FMapeamento.Values['CLASSIFICACAO_FISCAL'] <> '' then
              codCFI := Trim(dbfProdutos.FieldByName(FMapeamento.Values['CLASSIFICACAO_FISCAL']).AsString);

            desc := PrepareStringForDB(SanitizeText(descRaw));
            codBarras := PrepareStringForDB(SanitizeText(codBarrasRaw));
            un := PrepareStringForDB(SanitizeText(unRaw));
            ncm := PrepareStringForDB(SanitizeText(ncmRaw));
            grupo := PrepareStringForDB(SanitizeText(grupoRaw));
            marca := PrepareStringForDB(SanitizeText(marcaRaw));

            if un = '' then un := 'UN';
            codBarras := EnsureUniqueBarcode(query, codBarras);

            codGrupo := GetOrCreateGrupo(trans, grupo);
            codMarca := GetOrCreateMarca(trans, marca);

            idICMS := 0;
            if codCFI <> '' then
            begin
              if FICMSConfigCache.Values[codCFI] <> '' then
                idICMS := StrToIntDef(FICMSConfigCache.Values[codCFI], 0)
              else
                idICMS := BuscarICMSNoBanco(codCFI);
            end;

            if FUsarGrade then
              ImportarProdutoGrade(trans, desc, codBarras, un, ncm, precoVenda,
                precoCusto, precoEntrada, qtdEstoque, codGrupo, codMarca, idICMS, codBarrasAlt)
            else
              ImportarProdutoNormal(trans, desc, codBarras, un, ncm, precoVenda,
                precoCusto, precoEntrada, qtdEstoque, codGrupo, codMarca, idICMS, codBarrasAlt);

          except
            on E: Exception do
            begin
              LogErro(numLinha, E.Message, descRaw);
            end;
          end;

          if (atual mod BATCH_SIZE) = 0 then
            trans.CommitRetaining;

          dbfProdutos.Next;
        end;

        try
          query.SQL.Text := 'UPDATE PRODUTO SET ALTERADO = ' + QuotedStr(FormatDateTime('yyyy-mm-dd hh:nn:ss', Now));
          query.ExecSQL;
        except
          on E: Exception do
            LogMsg('log_erros.txt', 'Erro atualização ALTERADO: ' + E.Message);
        end;

        trans.Commit;
        t1 := Now;

        DoProgress(100);

        if FLinhasComErro.Count > 0 then
          DoProgressLabel(Format('✅ Concluído com %d erros em %d s (%d itens)', [FLinhasComErro.Count, SecondsBetween(t1, t0), totalGeral]))
        else
          DoProgressLabel(Format('✅ Concluído em %d s (%d itens)', [SecondsBetween(t1, t0), totalGeral]));

      finally
        if Assigned(dbfProdutos) then
        begin
          if dbfProdutos.Active then dbfProdutos.Close;
          FreeAndNil(dbfProdutos);
        end;
      end;

    finally
      selectedFields.Free;
    end;

  finally
    FreeAndNil(FQInsSimples);
    FreeAndNil(FQInsPai);
    FreeAndNil(FQInsFilho);
    FreeAndNil(FQInsGrade);
    FreeAndNil(FQInsVar);
    FreeAndNil(FQInsGrupo);
    FreeAndNil(FQInsMarca);
    FreeAndNil(FQInsICMS);
    FreeAndNil(query);
    FreeAndNil(trans);
    FreeAndNil(conn);
  end;
end;

procedure TImportadorBase.ImportarSomenteGrupos;
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfProdutos: TDbf;
  gruposUnicos: TStringList;
  descGrupo: string;
  total, atual: Integer;
begin
  if FMapeamento.Values['GRUPO'] = '' then
  begin
    DoProgressLabel('Erro: Coluna Grupo não mapeada');
    Exit;
  end;

  DoProgressLabel('Importando somente grupos...');
  DoProgress(0);

  gruposUnicos := TStringList.Create;
  gruposUnicos.Sorted := True;
  gruposUnicos.Duplicates := dupIgnore;

  conn := TIBConnection.Create(nil);
  trans := TSQLTransaction.Create(nil);
  query := TSQLQuery.Create(nil);
  dbfProdutos := TDbf.Create(nil);

  try
    try
      conn.DatabaseName := FCaminhoFDB;
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

      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D FROM GRUPO';
      query.Open;
      while not query.EOF do
      begin
        gruposUnicos.Add(query.FieldByName('D').AsString);
        query.Next;
      end;
      query.Close;

      dbfProdutos.FilePathFull := ExtractFilePath(FCaminhoDBF);
      dbfProdutos.TableName := ExtractFileName(FCaminhoDBF);
      dbfProdutos.Open;

      total := dbfProdutos.RecordCount;
      atual := 0;

      trans.StartTransaction;

      dbfProdutos.First;
      while not dbfProdutos.EOF do
      begin
        Inc(atual);
        descGrupo := Trim(dbfProdutos.FieldByName(FMapeamento.Values['GRUPO']).AsString);

        if (descGrupo <> '') and (gruposUnicos.IndexOf(UpperCase(descGrupo)) = -1) then
        begin
          FLastCodGrupo := NextCodigoGrupo;

          query.SQL.Text := 'INSERT INTO GRUPO (CODGRUPO, DESCRICAO) VALUES (:COD, :DESC)';
          query.ParamByName('COD').AsInteger := FLastCodGrupo;
          query.ParamByName('DESC').AsString := Copy(descGrupo, 1, 50);
          query.ExecSQL;

          gruposUnicos.Add(UpperCase(descGrupo));
        end;

        if (atual mod 100) = 0 then
        begin
          DoProgress(Round((atual / total) * 100));
          DoProgressLabel(Format('Grupos: %d/%d', [atual, total]));
        end;

        dbfProdutos.Next;
      end;

      trans.Commit;

      DoProgress(100);
      DoProgressLabel(Format('✅ Grupos importados: %d novos', [gruposUnicos.Count]));

    except
      on E: Exception do
      begin
        if trans.Active then
          trans.Rollback;
        DoProgressLabel('❌ Erro ao importar grupos: ' + E.Message);
        LogMsg('log_erros.txt', 'Erro importar grupos: ' + E.Message);
      end;
    end;

  finally
    dbfProdutos.Free;
    query.Free;
    trans.Free;
    conn.Free;
    gruposUnicos.Free;
  end;
end;

procedure TImportadorBase.ImportarSomenteMarcas;
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfProdutos: TDbf;
  marcasUnicas: TStringList;
  descMarca: string;
  total, atual: Integer;
begin
  if FMapeamento.Values['MARCA'] = '' then
  begin
    DoProgressLabel('Erro: Coluna Marca não mapeada');
    Exit;
  end;

  DoProgressLabel('Importando somente marcas...');
  DoProgress(0);

  marcasUnicas := TStringList.Create;
  marcasUnicas.Sorted := True;
  marcasUnicas.Duplicates := dupIgnore;

  conn := TIBConnection.Create(nil);
  trans := TSQLTransaction.Create(nil);
  query := TSQLQuery.Create(nil);
  dbfProdutos := TDbf.Create(nil);

  try
    try
      conn.DatabaseName := FCaminhoFDB;
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

      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D FROM MARCA';
      query.Open;
      while not query.EOF do
      begin
        marcasUnicas.Add(query.FieldByName('D').AsString);
        query.Next;
      end;
      query.Close;

      dbfProdutos.FilePathFull := ExtractFilePath(FCaminhoDBF);
      dbfProdutos.TableName := ExtractFileName(FCaminhoDBF);
      dbfProdutos.Open;

      total := dbfProdutos.RecordCount;
      atual := 0;

      trans.StartTransaction;

      dbfProdutos.First;
      while not dbfProdutos.EOF do
      begin
        Inc(atual);
        descMarca := Trim(dbfProdutos.FieldByName(FMapeamento.Values['MARCA']).AsString);

        if (descMarca <> '') and (marcasUnicas.IndexOf(UpperCase(descMarca)) = -1) then
        begin
          FLastCodMarca := NextCodigoMarca;

          query.SQL.Text := 'INSERT INTO MARCA (CODMARCA, DESCRICAO) VALUES (:COD, :DESC)';
          query.ParamByName('COD').AsInteger := FLastCodMarca;
          query.ParamByName('DESC').AsString := Copy(descMarca, 1, 50);
          query.ExecSQL;

          marcasUnicas.Add(UpperCase(descMarca));
        end;

        if (atual mod 100) = 0 then
        begin
          DoProgress(Round((atual / total) * 100));
          DoProgressLabel(Format('Marcas: %d/%d', [atual, total]));
        end;

        dbfProdutos.Next;
      end;

      trans.Commit;

      DoProgress(100);
      DoProgressLabel(Format('✅ Marcas importadas: %d novas', [marcasUnicas.Count]));

    except
      on E: Exception do
      begin
        if trans.Active then
          trans.Rollback;
        DoProgressLabel('❌ Erro ao importar marcas: ' + E.Message);
        LogMsg('log_erros.txt', 'Erro importar marcas: ' + E.Message);
      end;
    end;

  finally
    dbfProdutos.Free;
    query.Free;
    trans.Free;
    conn.Free;
    marcasUnicas.Free;
  end;
end;

procedure TImportadorBase.ImportarSomenteTributacao;
var
  conn: TIBConnection;
  trans: TSQLTransaction;
  query: TSQLQuery;
  dbfProdutos: TDbf;
  configsUnicas: TStringList;
  configTrib: string;
  total, atual: Integer;
begin
  DoProgressLabel('Importando configurações de tributação...');
  DoProgress(0);

  configsUnicas := TStringList.Create;
  configsUnicas.Sorted := True;
  configsUnicas.Duplicates := dupIgnore;

  conn := TIBConnection.Create(nil);
  trans := TSQLTransaction.Create(nil);
  query := TSQLQuery.Create(nil);
  dbfProdutos := TDbf.Create(nil);

  try
    try
      conn.DatabaseName := FCaminhoFDB;
      conn.UserName := 'sysdba';
      conn.Password := 'masterkey';
      conn.CharSet := 'UTF8';
      conn.Params.Add('sql_dialect=3');
      conn.Transaction := trans;

      trans.DataBase := conn;
      query.DataBase := conn;
      query.Transaction := trans;

      conn.Open;

      if not TableExists(query, 'ICMS') then
      begin
        DoProgressLabel('Tabela ICMS não encontrada no banco de dados.');
        Exit;
      end;

      query.SQL.Text := 'SELECT UPPER(DESCRICAO) AS D FROM ICMS';
      query.Open;
      while not query.EOF do
      begin
        configsUnicas.Add(query.FieldByName('D').AsString);
        query.Next;
      end;
      query.Close;

      dbfProdutos.FilePathFull := ExtractFilePath(FCaminhoDBF);
      dbfProdutos.TableName := ExtractFileName(FCaminhoDBF);
      dbfProdutos.Open;

      total := dbfProdutos.RecordCount;
      atual := 0;

      if trans.Active then
      trans.Commit;

      trans.StartTransaction;

      dbfProdutos.First;
      while not dbfProdutos.EOF do
      begin
        Inc(atual);

        if FMapeamento.Values['NCM'] <> '' then
        begin
          configTrib := 'NCM ' + Trim(dbfProdutos.FieldByName(FMapeamento.Values['NCM']).AsString);

          if (configTrib <> '') and (configsUnicas.IndexOf(UpperCase(configTrib)) = -1) then
          begin
            FLastID_ICMS := NextID_ICMS;

            query.SQL.Text :=
              'INSERT INTO ICMS (ID_ICMS, DESCRICAO, CST, MODBC, BASEICMS, ALIQICMS, ORIGEM) ' +
              'VALUES (:ID, :DS, ''00'', 3, 100.0, :ALIQ, 0)';
            query.ParamByName('ID').AsInteger := FLastID_ICMS;
            query.ParamByName('DS').AsString := Copy(configTrib, 1, 50);
            query.ParamByName('ALIQ').AsFloat := 17.0;
            query.ExecSQL;

            configsUnicas.Add(UpperCase(configTrib));
          end;
        end;

        if (atual mod 100) = 0 then
        begin
          DoProgress(Round((atual / total) * 100));
          DoProgressLabel(Format('Tributação: %d/%d', [atual, total]));
        end;

        dbfProdutos.Next;
      end;

      trans.Commit;

      DoProgress(100);
      DoProgressLabel(Format('✅ Configurações de tributação importadas: %d novas', [configsUnicas.Count]));

    except
      on E: Exception do
      begin
        if trans.Active then
          trans.Rollback;
        DoProgressLabel('❌ Erro ao importar tributação: ' + E.Message);
        LogMsg('log_erros.txt', 'Erro importar tributação: ' + E.Message);
      end;
    end;

  finally
    dbfProdutos.Free;
    query.Free;
    trans.Free;
    conn.Free;
    configsUnicas.Free;
  end;
end;

end.
