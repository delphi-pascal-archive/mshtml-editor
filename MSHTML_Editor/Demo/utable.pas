unit utable;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons;

type
  TfrmTable = class(TForm)
    Button1: TButton;
    Button2: TButton;
    Table: TGroupBox;
    Cells: TGroupBox;
    Label1: TLabel;
    Label2: TLabel;
    lblWidth: TLabel;
    edRows: TEdit;
    edCols: TEdit;
    edWidth: TEdit;
    edTBorder: TEdit;
    Label3: TLabel;
    Label5: TLabel;
    Label6: TLabel;
    chkWrap: TCheckBox;
    cmbTableBGColor: TComboBox;
    edCaption: TEdit;
    edPad: TEdit;
    edSpace: TEdit;
    Label4: TLabel;
    Label7: TLabel;
    ColorDialog1: TColorDialog;
    btnTableBGColor: TSpeedButton;
    procedure Button2Click(Sender: TObject);
    procedure Button1Click(Sender: TObject);
    procedure btnTableBGColorClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmTable: TfrmTable;

implementation

{$R *.DFM}

//------------------------------------------------------------------------------
function Color2HTML(Color: TColor): string;
var //9-5-99 New RuudK
  Step  : Integer;
  Step2 : String;
begin
  Step := ColorToRGB(Color);
  Step2 := IntToHex(Step, 6);
  //The result will be BBGGRR but I want it RRGGBB
  Result := '#' + Copy(Step2, 5, 2) + Copy(Step2, 3, 2) + Copy(Step2, 1, 2);
end;

//------------------------------------------------------------------------------
procedure TfrmTable.Button2Click(Sender: TObject);
begin
  ModalResult := mrOK;
end;

//------------------------------------------------------------------------------
procedure TfrmTable.Button1Click(Sender: TObject);
begin
  Close;
end;

//------------------------------------------------------------------------------
procedure TfrmTable.btnTableBGColorClick(Sender: TObject);
//Added 16-5-99 RuudK
var
  tableBGColor: string;
begin
  if ColorDialog1.Execute then begin
    tableBGColor := Color2HTML(ColorDialog1.Color);
    cmbTableBGColor.Text := tableBGColor;
  end;
end;


end.
