unit uPageProp;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, Jpeg, ExtDlgs, ExtCtrls;

type
  TfrmPageProp = class(TForm)
    GroupBox1: TGroupBox;
    cmbBackColor: TComboBox;
    GroupBox2: TGroupBox;
    GroupBox3: TGroupBox;
    edTitle: TEdit;
    btnLoad: TSpeedButton;
    btnOK: TButton;
    btnCancel: TButton;
    OpenPictureDialog: TOpenPictureDialog;
    ColorDialog1: TColorDialog;
    btnBGColor: TSpeedButton;
    Image1: TImage;
    GroupBox4: TGroupBox;
    btnLColor: TSpeedButton;
    cmbLinkColor: TComboBox;
    GroupBox5: TGroupBox;
    btnVLColor: TSpeedButton;
    cmbVLinkColor: TComboBox;
    GroupBox6: TGroupBox;
    btnALColor: TSpeedButton;
    cmbALinkColor: TComboBox;
    procedure btnLoadClick(Sender: TObject);
    procedure btnBGColorClick(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure btnLColorClick(Sender: TObject);
    procedure btnVLColorClick(Sender: TObject);
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmPageProp: TfrmPageProp;

implementation

uses Main;

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
procedure TfrmPageProp.btnLoadClick(Sender: TObject);
begin
  if OpenPictureDialog.Execute then
   Image1.Picture.LoadFromFile(OpenPictureDialog.Filename);
end;

//------------------------------------------------------------------------------
procedure TfrmPageProp.btnBGColorClick(Sender: TObject);
//Added 15-5-99 RuudK
var
  docALColor: string;
begin
  if ColorDialog1.Execute then begin
    docALColor := Color2HTML(ColorDialog1.Color);
    cmbALinkColor.Text := docALColor;
  end;
end;

//------------------------------------------------------------------------------
procedure TfrmPageProp.FormShow(Sender: TObject);
begin
  //edTitle.Text:=frmMain.DHTMLEdit.DocumentTitle;
end;

procedure TfrmPageProp.btnLColorClick(Sender: TObject);
//Added 15-5-99 RuudK
var
  docLColor: string;
begin
  if ColorDialog1.Execute then begin
    docLColor := Color2HTML(ColorDialog1.Color);
    cmbLinkColor.Text := docLColor;
  end;
end;  

procedure TfrmPageProp.btnVLColorClick(Sender: TObject);
//Added 15-5-99 RuudK
var
  docVLColor: string;
begin
  if ColorDialog1.Execute then begin
    docVLColor := Color2HTML(ColorDialog1.Color);
    cmbVLinkColor.Text := docVLColor;
  end;
end;

end.

