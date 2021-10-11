unit ugrid;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls;

type
  TfrmGrid = class(TForm)
    edX: TEdit;
    edY: TEdit;
    Label1: TLabel;
    Label2: TLabel;
    Button1: TButton;
    Bevel1: TBevel;
    CheckBox1: TCheckBox;
    procedure edYChange(Sender: TObject);
    procedure edXChange(Sender: TObject);
    procedure CheckBox1Click(Sender: TObject);
    procedure FormCreate(Sender: TObject);

  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmGrid: TfrmGrid;


procedure DoDlg_Grid;

implementation

uses main;

{$R *.DFM}

//------------------------------------------------------------------------------
procedure DoDlg_Grid;
begin

        frmGrid := TfrmGrid.Create(Nil);
        try

           frmGrid.ShowModal;
        finally

           frmGrid.Free;
        end;

end;
(*
//------------------------------------------------------------------------------
procedure TfrmGrid.FormCreate(Sender: TObject);
begin
   edX.Text := IntToStr(aEditHost.GridX);
   edY.Text := IntToStr(aEditHost.GridY);
   CheckBox1.checked := aEditHost.Snap;
end;
//------------------------------------------------------------------------------
procedure TfrmGrid.edXChange(Sender: TObject);
begin
  aEditHost.GridX :=StrToInt(edX.Text);
end;
//------------------------------------------------------------------------------
procedure TfrmGrid.edYChange(Sender: TObject);
begin
  aEditHost.GridY :=StrToInt(edY.Text);
end;
//------------------------------------------------------------------------------
procedure TfrmGrid.CheckBox1Click(Sender: TObject);
begin
  aEditHost.Snap := CheckBox1.checked;
end;
//------------------------------------------------------------------------------
*)



procedure TfrmGrid.edYChange(Sender: TObject);
begin
  //frmMain.DHTMLEdit.SnapToGridX :=StrToInt(edY.Text);
  frmMain.Edit.SnapToGridY :=StrToInt(edY.Text);
end;

procedure TfrmGrid.edXChange(Sender: TObject);
begin
  //frmMain.DHTMLEdit.SnapToGridX :=StrToInt(edX.Text);
  frmMain.Edit.SnapToGridX :=StrToInt(edX.Text);
end;

procedure TfrmGrid.CheckBox1Click(Sender: TObject);
begin
  //frmMain.DHTMLEdit.SnapToGrid;
  frmMain.Edit.SnapToGrid := CheckBox1.checked;
end;

procedure TfrmGrid.FormCreate(Sender: TObject);
begin
   edX.Text := IntToStr(frmMain.Edit.SnapToGridX);
   edY.Text := IntToStr(frmMain.Edit.SnapToGridY);
   CheckBox1.checked := frmMain.Edit.SnapToGrid;

   //edX.Text := IntToStr(SnapToGridX);
   //edY.Text := IntToStr(SnapToGridY);
   //CheckBox1.checked := SnapToGrid;
end;


end.
