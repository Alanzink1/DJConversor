unit uImportacaoMarca;

{$mode objfpc}{$H+}

interface

uses
  Classes, SysUtils, DB, dbf, IBConnection, SQLDB, Forms, Controls, Graphics,
  Dialogs, StdCtrls, ComCtrls, ExtCtrls, Grids, Buttons, StrUtils, DateUtils,
  Variants, uImportadorBase;

type
  TForm6 = class(TForm)
    PanelTop: TPanel;
    LabelTitulo: TLabel;
    PanelBotoes: TPanel;
    ButtonImportar: TButton;
    ButtonFechar: TButton;
    PanelConfig: TPanel;
    LabelArquivo: TLabel;
    EditCaminhoDBF: TEdit;
    ButtonBuscarDBF: TButton;
    LabelBanco: TLabel;
    EditCaminhoFDB: TEdit;
    ButtonBuscarFDB: TButton;
    LabelColuna: TLabel;
    ComboColunaMarca: TComboBox;
    OpenDialog: TOpenDialog;
    ProgressBar: TProgressBar;
    LabelProgresso: TLabel;
    StringGridMarcas: TStringGrid;
    PanelGrid: TPanel;
    LabelGrid: TLabel;
    ButtonCarregar: TButton;
    CheckBoxSobrescrever: TCheckBox;
    CheckBoxApenasNovas: TCheckBox;

    procedure ButtonBuscarDBFClick(Sender: TObject);
    procedure ButtonBuscarFDBClick(Sender: TObject);
    procedure ButtonImportarClick(Sender: TObject);
    procedure ButtonFecharClick(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure FormClose(Sender: TObject; var CloseAction: TCloseAction);
    procedure ButtonCarregarClick(Sender: TObject);
    procedure ComboColunaMarcaChange(Sender: TObject);

  private

  public
  end;

var
  Form6: TForm6;

implementation

{$R *.lfm}

{ TForm6 }


end.
