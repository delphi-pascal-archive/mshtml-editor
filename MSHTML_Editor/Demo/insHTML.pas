unit insHTML;

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, ExtCtrls;

type
  TfrmInsertHTML = class(TForm)
    Label1: TLabel;
    Button1: TButton;
    Button2: TButton;
    InsHTML: TMemo;
    Bevel1: TBevel;
  private
    { Private declarations }
  public
    { Public declarations }
  end;

var
  frmInsertHTML: TfrmInsertHTML;

implementation

{$R *.DFM}

end.
