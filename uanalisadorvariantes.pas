unit uAnalisadorVariantes;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, RegExpr;

type
  TAnalisadorVariantes = class
  private
    FDescricao: string;
    FPalavras: TStringList;
    FBaseProduto: string;
    FVariacoes: TStringList;

    procedure TokenizarDescricao;
    function IsPalavraCor(const palavra: string): Boolean;
    function IsPalavraTamanho(const palavra: string): Boolean;
    function IsPadraoTamanho(const palavra: string): Boolean;
    function ExtrairBaseProduto: string;

  public
    constructor Create(const ADescricao: string);
    destructor Destroy; override;

    function TemVariacao: Boolean;
    function GetTipoVariacao: string;
    function GetBaseProduto: string;
    function GetVariacoes: TStringList;
    function GetVariacaoPrincipal: string;
  end;

const
  PALAVRAS_CHAVE_COR: array[0..20] of string = (
    'COR', 'COLOR', 'COLORIDO', 'COLORIDA', 'TINTA', 'PIGMENTO',
    'VERMELHO', 'VERDE', 'AZUL', 'AMARELO', 'PRETO', 'BRANCO', 'CINZA',
    'ROSA', 'ROXO', 'LARANJA', 'MARROM', 'BEGE', 'DOURADO', 'PRATEADO', 'TRANSPARENTE'
  );

  PALAVRAS_CHAVE_TAMANHO: array[0..15] of string = (
    'TAM', 'TAMANHO', 'SIZE', 'PORTE', 'MEDIDA', 'DIMENSAO',
    'P', 'M', 'G', 'GG', 'XG', 'PP', 'MG', 'UNICO', 'ÚNICO', 'UNITARIO'
  );

  PADROES_TAMANHO: array[0..8] of string = (
    '^\d+[Xx]\d+$', '^\d+[ML]$', '^\d+[CM]$', '^\d+(\.\d+)?[ML]$',
    '^[PMGX]+$', '^\d+$', '^\d+/\d+$', '^[0-9\-]+$', '^UNICO|ÚNICO$'
  );

implementation

{ TAnalisadorVariantes }

constructor TAnalisadorVariantes.Create(const ADescricao: string);
begin
  FDescricao := UpperCase(Trim(ADescricao));
  FPalavras := TStringList.Create;
  FVariacoes := TStringList.Create;
  FVariacoes.Duplicates := dupIgnore;
  FVariacoes.Sorted := True;

  TokenizarDescricao;
  FBaseProduto := ExtrairBaseProduto;
end;

destructor TAnalisadorVariantes.Destroy;
begin
  FPalavras.Free;
  FVariacoes.Free;
  inherited Destroy;
end;

procedure TAnalisadorVariantes.TokenizarDescricao;
var
  i: Integer;
  palavra: string;
  palavrasTemp: TStringList;
begin
  palavrasTemp := TStringList.Create;
  try
    palavrasTemp.Delimiter := ' ';
    palavrasTemp.DelimitedText := StringReplace(FDescricao, '-', ' ', [rfReplaceAll]);
    palavrasTemp.DelimitedText := StringReplace(palavrasTemp.DelimitedText, '/', ' ', [rfReplaceAll]);
    palavrasTemp.DelimitedText := StringReplace(palavrasTemp.DelimitedText, '_', ' ', [rfReplaceAll]);

    for i := 0 to palavrasTemp.Count - 1 do
    begin
      palavra := Trim(palavrasTemp[i]);
      if palavra <> '' then
        FPalavras.Add(palavra);
    end;
  finally
    palavrasTemp.Free;
  end;
end;

function TAnalisadorVariantes.IsPalavraCor(const palavra: string): Boolean;
var
  i: Integer;
  pal: string;
begin
  pal := UpperCase(Trim(palavra));
  Result := False;

  for i := Low(PALAVRAS_CHAVE_COR) to High(PALAVRAS_CHAVE_COR) do
  begin
    if pal = PALAVRAS_CHAVE_COR[i] then
    begin
      Result := True;
      Exit;
    end;
  end;

  if (Length(pal) >= 3) and
     (Pos('AZUL', pal) = 1) or (Pos('VERM', pal) = 1) or
     (Pos('VERD', pal) = 1) or (Pos('AMAR', pal) = 1) or
     (Pos('PRET', pal) = 1) or (Pos('BRAN', pal) = 1) or
     (Pos('ROSA', pal) = 1) or (Pos('ROXO', pal) = 1) or
     (Pos('CINZ', pal) = 1) or (Pos('MAR', pal) = 1) then
  begin
    Result := True;
  end;
end;

function TAnalisadorVariantes.IsPalavraTamanho(const palavra: string): Boolean;
var
  i: Integer;
  pal: string;
begin
  pal := UpperCase(Trim(palavra));
  Result := False;

  for i := Low(PALAVRAS_CHAVE_TAMANHO) to High(PALAVRAS_CHAVE_TAMANHO) do
  begin
    if pal = PALAVRAS_CHAVE_TAMANHO[i] then
    begin
      Result := True;
      Exit;
    end;
  end;

  Result := IsPadraoTamanho(pal);
end;

function TAnalisadorVariantes.IsPadraoTamanho(const palavra: string): Boolean;
var
  i: Integer;
  regex: TRegExpr;
  pal: string;
begin
  pal := UpperCase(Trim(palavra));
  Result := False;

  for i := Low(PADROES_TAMANHO) to High(PADROES_TAMANHO) do
  begin
    regex := TRegExpr.Create;
    try
      regex.Expression := PADROES_TAMANHO[i];
      if regex.Exec(pal) then
      begin
        Result := True;
        Break;
      end;
    finally
      regex.Free;
    end;
  end;
end;

function TAnalisadorVariantes.ExtrairBaseProduto: string;
var
  i: Integer;
  baseParts: TStringList;
  palavra: string;
  encontrouVariacao: Boolean;
begin
  baseParts := TStringList.Create;
  encontrouVariacao := False;

  try
    for i := 0 to FPalavras.Count - 1 do
    begin
      palavra := FPalavras[i];

      if IsPalavraCor(palavra) or IsPalavraTamanho(palavra) then
      begin
        encontrouVariacao := True;
        FVariacoes.Add(palavra);
      end
      else if not encontrouVariacao then
      begin
        baseParts.Add(palavra);
      end
      else
      begin
        if (i > 0) and (baseParts.Count > 0) then
        begin
          baseParts.Add(palavra);
        end;
      end;
    end;

    Result := baseParts.DelimitedText;
    Result := StringReplace(Result, '"', '', [rfReplaceAll]);
    Result := Trim(Result);
  finally
    baseParts.Free;
  end;
end;

function TAnalisadorVariantes.TemVariacao: Boolean;
begin
  Result := FVariacoes.Count > 0;
end;

function TAnalisadorVariantes.GetTipoVariacao: string;
var
  temCor, temTamanho: Boolean;
  i: Integer;
begin
  temCor := False;
  temTamanho := False;

  for i := 0 to FVariacoes.Count - 1 do
  begin
    if IsPalavraCor(FVariacoes[i]) then
      temCor := True
    else if IsPalavraTamanho(FVariacoes[i]) then
      temTamanho := True;
  end;

  if temCor and temTamanho then
    Result := 'AMBOS'
  else if temCor then
    Result := 'COR'
  else if temTamanho then
    Result := 'TAMANHO'
  else
    Result := 'NENHUM';
end;

function TAnalisadorVariantes.GetBaseProduto: string;
begin
  Result := FBaseProduto;
end;

function TAnalisadorVariantes.GetVariacoes: TStringList;
begin
  Result := FVariacoes;
end;

function TAnalisadorVariantes.GetVariacaoPrincipal: string;
begin
  if FVariacoes.Count > 0 then
    Result := FVariacoes[0]
  else
    Result := '';
end;

end.
