program DHTML;

uses
  Forms,
  main in 'main.pas' {frmMain},
  utable in 'utable.pas' {frmTable},
  uPageProp in 'uPageProp.pas' {frmPageProp},
  ugrid in 'ugrid.pas' {frmGrid},
  insHTML in 'insHTML.pas' {frmInsertHTML},
  FmxUtils in 'Fmxutils.pas',
  SysUtils,
  Windows;

{$R *.RES}


begin
  Application.Initialize;
  Application.Title := 'MSHTML Editor';
  Application.CreateForm(TfrmMain, frmMain);
  Application.Run;
end.
