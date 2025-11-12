unit uCompartilhado;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, SQLDB, DB;


function RemoveAcentos(const S: string): string;
function SanitizeText(const S: string): string;
function PrepareStringForDB(const S: string): string;
procedure LogMsg(const nomeArq, msg: string);
function ParseDateBR(const S: string): TDateTime;
function ParseFloatBR(const S: string): Double;
procedure Prep(Q: TSQLQuery; const SQLText: string);

implementation

uses
  DateUtils;

const
  LOGS_ENABLED = True;

function RemoveAcentos(const S: string): string;
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

function SanitizeText(const S: string): string;
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

function PrepareStringForDB(const S: string): string;
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

procedure LogMsg(const nomeArq, msg: string);
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

function ParseDateBR(const S: string): TDateTime;
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

function ParseFloatBR(const S: string): Double;
var
  T: string;
begin
  T := Trim(S);
  T := StringReplace(T, '.', '', [rfReplaceAll]);
  T := StringReplace(T, ',', '.', [rfReplaceAll]);
  if not TryStrToFloat(T, Result) then
    Result := 0;
end;

procedure Prep(Q: TSQLQuery; const SQLText: string);
begin
  Q.Close;
  if Q.Prepared then
    Q.UnPrepare;
  Q.SQL.Text := SQLText;
  Q.Prepare;
end;

end.
