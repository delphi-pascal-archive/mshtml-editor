{ Open source project MSHTML Editor OSPv3r200 05-01-2004

  This editor is build from a DHTML edit OSP from 16-05-1999 found at:
  delphi-dhtmledit@yahoogroups.com
  http://groups.yahoo.com/group/delphi-dhtmledit/files/DHTMLEdit/DHTML.zip
  (look in the file section)

  This update editor is created by Kurt Senfer
  "Copyright (C) 2002-2004  Kurt Senfer <Kurt.Senfer@siemens.com>"

  The code in this unit is released under:
  GNU General Public License - http://www.gnu.org/copyleft/gpl.html

  Exept fore the original source as it is implemented in the original OSP,
  witch have the same copyright status as found in the original package DHTML.zip.

  I have often, in various mails on the delphi-dhtmledit list, pointed out that
  The DHTMLEdit component is just a wrapper around the MSHTML engine - the
  Wrapper contains coding that enables additional features like:
     Table operations (insert row, insert column, merge cells, and so on).
     Absolute drop mode.
     Source code white space and formatting preservation.
     Custom Editing Glyphs.
     ........


  Replacing DHTMLEdit with the MSHTML engine (or as in this case with the
  EmbeddedED component) means that you need to code a lot of features, like
  tablehandling and so on, witch is build in fearture in DHTMLEdit, yourself.

  I chosed to create a VCL wrapper around MSHTML and named it:
  TEmbeddedED~ Embedded EDitor
  TEmbeddedED implements most of the code you need to replace the DHTML Component
  (if you only need a basic editor) and it offers a lot of features never
  found in DHTML Edit. The main reason for encapsulating all that code in a
  component is that its easy to update existing editor projects with
  new features (in new versions of TEmbeddedED).


  This is a short listing of the main things you have to do in order to skip
  DHTMLEdit in favour of the HMHTML engine in your own DHTMLEdit project.

  - Replace TDHTMLEdit with MSHTML
  - Call the existing DHTMLEditDisplayChanged from a new "OnUpdateUI function"
  - Implement Custom Editing Glyphs
  - Implement IHTMLEditDesigner
  - Implement IHTMLEditHost
  - Implemet code to handle the table stuff
  - Implemet code to handle the position stuff
}


unit main;

{$I KSED.INC} //Compiler version directives

interface

uses
  Windows, Messages, SysUtils, Classes, Graphics, Controls, Forms, Dialogs,
  StdCtrls, Buttons, ComCtrls, Menus, ExtCtrls, ActiveX, mshtml_tlb, FmxUtils,
  ImgList,  OleCtrls, EmbeddedED, ToolWin, SHDocVw;  // SHDocVw_TLB, 

type
  CmdID = TOleEnum;

  TGetInputArg = procedure(Sender: TObject; var pVarIn: OleVariant; var FuncResult: boolean) of object;
  TClearInputArg = procedure(Sender: TObject; var pVarIn: OleVariant) of object;
  TfrmMain = class(TForm)
    Panel1: TPanel;
    StatusBar1: TStatusBar;
    MainMenu1: TMainMenu;
    mnuFile: TMenuItem;
    mnuFileNew: TMenuItem;
    mnuFileOpen: TMenuItem;
    mnuFileSave: TMenuItem;
    mnuFileSaveAs: TMenuItem;
    N2: TMenuItem;
    mnuFilePrint: TMenuItem;
    mnuFilePageSetup: TMenuItem;
    N1: TMenuItem;
    mnuFileExit: TMenuItem;
    mnuInsert: TMenuItem;
    mnuTableInsertTable: TMenuItem;
    mnuInsertImage: TMenuItem;
    PopupMenu1: TPopupMenu;
    mnuEdit: TMenuItem;
    mnuEditDelete: TMenuItem;
    mnuEditCut: TMenuItem;
    mnuEditCopy: TMenuItem;
    mnuEditPaste: TMenuItem;
    mnuInsertHyperlink: TMenuItem;
    N3: TMenuItem;
    mnuEditUndo: TMenuItem;
    mnuEditRedo: TMenuItem;
    N4: TMenuItem;
    mnuEditSelectAll: TMenuItem;
    N5: TMenuItem;
    mnuEditUnlink: TMenuItem;
    mnuTable: TMenuItem;
    mnuTableInsertRow: TMenuItem;
    mnuTableInsertCol: TMenuItem;
    mnuTableInsetCell: TMenuItem;
    mnuTableDeleteRow: TMenuItem;
    mnuTableDeleteCol: TMenuItem;
    mnuTableDeleteCell: TMenuItem;
    N6: TMenuItem;
    mnuTableMergeCells: TMenuItem;
    mnuTableSplitCells: TMenuItem;
    ColorDialog: TColorDialog;
    mnuPosition: TMenuItem;
    mnuPositionAboveText: TMenuItem;
    mnuPositionBringForward: TMenuItem;
    mnuPositionBringtoFront: TMenuItem;
    mnuPositionBelowText: TMenuItem;
    mnuPositionSendtoBack: TMenuItem;
    mnuPositionSendBackward: TMenuItem;
    N7: TMenuItem;
    pmuUndo: TMenuItem;
    pmuRedo: TMenuItem;
    pmuCopy: TMenuItem;
    pmuPaste: TMenuItem;
    pmuCut: TMenuItem;
    pmuDelete: TMenuItem;
    pmuSelectAll: TMenuItem;
    pmuUnlink: TMenuItem;
    pmuHyperlink: TMenuItem;
    mnuPositionMakeAbsolute: TMenuItem;
    N8: TMenuItem;
    mnuFileProperties: TMenuItem;
    mnuHelpAbout: TMenuItem;
    N9: TMenuItem;
    mnuInsertHTML: TMenuItem;
    mnuPositionGridSettings: TMenuItem;
    PageControl1: TPageControl;
    tabEdit: TTabSheet;
    tabHTML: TTabSheet;
    tabPreview: TTabSheet;
    mnuEditRemoveTags: TMenuItem;
    mmoHTML: TMemo;
    CoolBar: TCoolBar;
    tlb_Edit: TToolBar;
    btnNew: TToolButton;
    btnOpen: TToolButton;
    btnSave: TToolButton;
    ToolButton10: TToolButton;
    btnPrint: TToolButton;
    ToolButton9: TToolButton;
    btnCut: TToolButton;
    btnCopy: TToolButton;
    btnPaste: TToolButton;
    ToolButton1: TToolButton;
    btnUndo: TToolButton;
    btnRedo: TToolButton;
    tlb_Format: TToolBar;
    cmbStyles: TComboBox;
    ToolButton2: TToolButton;
    cmbFontName: TComboBox;
    ToolButton4: TToolButton;
    cmbFontSize: TComboBox;
    ToolButton6: TToolButton;
    btnFGColor: TToolButton;
    btnBGColor: TToolButton;
    ToolButton3: TToolButton;
    btnBold: TToolButton;
    btnItalic: TToolButton;
    btnUnderline: TToolButton;
    ToolButton7: TToolButton;
    btnAlignLeft: TToolButton;
    btnAlignCenter: TToolButton;
    btnAlignRight: TToolButton;
    ToolButton12: TToolButton;
    btnNumberedList: TToolButton;
    btnBulletedList: TToolButton;
    ToolButton15: TToolButton;
    btnIndentLeft: TToolButton;
    btnIndentRight: TToolButton;
    ImageList1: TImageList;
    ImageList2: TImageList;
    ToolButton11: TToolButton;
    btnInsertTable: TToolButton;
    ToolButton5: TToolButton;
    btnInsertLink: TToolButton;
    ToolButton14: TToolButton;
    btnInsertImage: TToolButton;
    ToolButton17: TToolButton;
    btnFind: TToolButton;
    mnuHelp: TMenuItem;
    N11: TMenuItem;
    N12: TMenuItem;
    N13: TMenuItem;
    N14: TMenuItem;
    mnuEditFind: TMenuItem;
    ToolButton8: TToolButton;
    btnMakeAbsolute: TToolButton;
    btnPrevInBrowser: TToolButton;
    ToolButton13: TToolButton;
    ImageList3: TImageList;
    ToolBar1: TToolBar;
    ToolBar2: TToolBar;
    tlb_Table: TToolBar;
    btnTableDelCol: TToolButton;
    btnTableInsCol: TToolButton;
    btnTableInsRow: TToolButton;
    btnTableDelRow: TToolButton;
    btnTableDelCell: TToolButton;
    btnTableInsCell: TToolButton;
    ToolButton23: TToolButton;
    ToolButton24: TToolButton;
    ToolButton25: TToolButton;
    btnTableInTable: TToolButton;
    btnTableMergeCells: TToolButton;
    ToolButton28: TToolButton;
    btnTableTableSplitCells: TToolButton;
    mnuView: TMenuItem;
    mnuToolbar: TMenuItem;
    mnuViewtbl_Table: TMenuItem;
    mnuViewtbl_Edit: TMenuItem;
    mnuViewtbl_Format: TMenuItem;
    btnShowDetails: TToolButton;
    btnBorders: TToolButton;
    btnLocalUndo: TToolButton;
    btnDropUndo: TToolButton;
    mnuPrintpreview: TMenuItem;
    PrintDialog1: TPrintDialog;
    btnJustify: TToolButton;
    btnFont: TToolButton;
    mnuLockElement: TMenuItem;
    N10: TMenuItem;
    mnuConstrainElement: TMenuItem;
    N15: TMenuItem;
    Edit: TEmbeddedED;
    Preview: TEmbeddedED;
    SpeedButton1: TSpeedButton;
    SpeedButton2: TSpeedButton;
    Memo1: TMemo;
    procedure EditDisplayChanged(Sender: TObject);
    procedure btnBoldClick(Sender: TObject);
    procedure btnOpenClick(Sender: TObject);
    procedure btnSaveClick(Sender: TObject);
    procedure btnInsertImageClick(Sender: TObject);
    procedure btnItalicClick(Sender: TObject);
    procedure btnUnderlineClick(Sender: TObject);
    procedure btnFindClick(Sender: TObject);
    procedure cmbFontNameChange(Sender: TObject);
    procedure cmbFontSizeChange(Sender: TObject);
    procedure cmbStylesChange(Sender: TObject);
    procedure mnuFileExitClick(Sender: TObject);
    procedure btnInsertTableClick(Sender: TObject);
    procedure mnuEditDeleteClick(Sender: TObject);
    procedure mnuEditCutClick(Sender: TObject);
    procedure mnuEditCopyClick(Sender: TObject);
    procedure mnuEditPasteClick(Sender: TObject);
    procedure mnuInsertHyperlinkClick(Sender: TObject);
    procedure btnInsertLinkClick(Sender: TObject);
    procedure btnAlignCenterClick(Sender: TObject);
    procedure btnAlignLeftClick(Sender: TObject);
    procedure btnAlignRightClick(Sender: TObject);
    procedure mnuEditRedoClick(Sender: TObject);
    procedure mnuEditUndoClick(Sender: TObject);
    procedure mnuEditSelectAllClick(Sender: TObject);
    procedure mnuEditUnlinkClick(Sender: TObject);
    procedure mnuTableInsertRowClick(Sender: TObject);
    procedure mnuTableInsertColClick(Sender: TObject);
    procedure mnuTableInsetCellClick(Sender: TObject);
    procedure mnuTableDeleteRowClick(Sender: TObject);
    procedure mnuTableDeleteColClick(Sender: TObject);
    procedure mnuTableDeleteCellClick(Sender: TObject);
    procedure btnFGColorClick(Sender: TObject);
    procedure mnuTableMergeCellsClick(Sender: TObject);
    procedure mnuTableSplitCellsClick(Sender: TObject);
    procedure btnIndentRightClick(Sender: TObject);
    procedure btnIndentLeftClick(Sender: TObject);
    procedure btnBulletedListClick(Sender: TObject);
    procedure btnNumberedListClick(Sender: TObject);
    procedure mnuPositionAboveTextClick(Sender: TObject);
    procedure mnuPositionBringForwardClick(Sender: TObject);
    procedure mnuPositionBringtoFrontClick(Sender: TObject);
    procedure mnuPositionBelowTextClick(Sender: TObject);
    procedure mnuPositionSendtoBackClick(Sender: TObject);
    procedure mnuPositionSendBackwardClick(Sender: TObject);
    procedure btnBGColorClick(Sender: TObject);
    procedure mnuPositionMakeAbsoluteClick(Sender: TObject);
    procedure mnuFilePropertiesClick(Sender: TObject);
    procedure btnPrintClick(Sender: TObject);
    procedure mnuHelpAboutClick(Sender: TObject);
    procedure mnuInsertHTMLClick(Sender: TObject);
    procedure btnNewClick(Sender: TObject);
    procedure mnuPositionGridSettingsClick(Sender: TObject);
    procedure mnuFileSaveAsClick(Sender: TObject);
    procedure mnuEditRemoveTagsClick(Sender: TObject);
    procedure PageControl1Changing(Sender: TObject; var AllowChange: Boolean);
    procedure btnPrevInBrowserClick(Sender: TObject);
    procedure mmoHTMLChange(Sender: TObject);
    procedure FormCreate(Sender: TObject);
    procedure mnuViewtbl_TableClick(Sender: TObject);
    procedure mnuViewtbl_FormatClick(Sender: TObject);
    procedure mnuViewtbl_EditClick(Sender: TObject);
    procedure EditDocumentComplete(Sender: TObject);
    procedure btnShowDetailsClick(Sender: TObject);
    procedure PageControl1Change(Sender: TObject);
    function EditQueryService(const rsid, iid: TGUID; out Obj: IUnknown): HRESULT;
    procedure btnBordersClick(Sender: TObject);
    procedure btnLocalUndoClick(Sender: TObject);
    procedure btnDropUndoClick(Sender: TObject);

    procedure tlb_FormatMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
    procedure cmbDropDown(Sender: TObject);
    procedure btnJustifyClick(Sender: TObject);
    procedure btnFontClick(Sender: TObject);
    procedure mnuEditClick(Sender: TObject);
    procedure mnuTableClick(Sender: TObject);
    procedure mnuPositionClick(Sender: TObject);
    procedure mnuLockElementClick(Sender: TObject);
    procedure mnuConstrainElementClick(Sender: TObject);
    procedure cmbFontSizeDropDown(Sender: TObject);
    procedure FormShow(Sender: TObject);
    procedure EditShowContextMenuEx(Sender: TObject; xPos, yPos: Integer);
    procedure EditPostEditorEventNotify(Sender: TObject;
      inEvtDispId: Integer; const pIEventObj: IHTMLEventObj;
      var aResult: HRESULT);
    procedure EditPreHandleEvent(Sender: TObject; inEvtDispId: Integer;
      const pIEventObj: IHTMLEventObj; var aResult: HRESULT);
    procedure EditInitialize(Sender: TObject; var NewFile: String);
    procedure mnuPrintpreviewClick(Sender: TObject);
    procedure FormClose(Sender: TObject; var Action: TCloseAction);
    procedure SpeedButton1Click(Sender: TObject);
    procedure SpeedButton2Click(Sender: TObject);

  private
    Changed: boolean;
    FFocusObject: TComboBox;
    DoPrintTime: Boolean;
    SourceView: Boolean;
    TEIgnoreChange: boolean;
    _CurrentDocumentPath: string;
    procedure LookForObjects;
    procedure Print(ShowDlg: boolean = True);
    procedure PrintHTMLSource;
    procedure UpdateButtons;
    procedure mnuQueryStatus(cmdID: Cardinal; var amnu: TMenuItem);
  end;




var
  frmMain: TfrmMain;

implementation

uses utable, insHTML, uPageProp, ugrid, Ieconst, IEDispConst, Printers, EmbedEDconst,
     PrintText, StdVCL, {$IFDEF D6D7} Variants, {$ENDIF} AxCtrls;

var
  LoadFirstFile: Boolean = True;

{$R *.DFM}


//------------------------------------------------------------------------------
function Color2HTML(Color: TColor): string;
var
  Step  : Integer;
  Step2 : String;
begin
  Step := ColorToRGB(Color);
  Step2 := IntToHex(Step, 6);
  //The result will be BBGGRR but I want it RRGGBB
  Result := '#' + Copy(Step2, 5, 2) + Copy(Step2, 3, 2) + Copy(Step2, 1, 2);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.EditDisplayChanged(Sender: TObject);
var
  Index: Integer;
  S: String;
begin
  //asm int 3 end; //trap

  if Edit.BrowseMode
     then exit;

  LookForObjects;

  UpdateButtons;

  //UpdateFontName
  cmbFontName.ItemIndex := Edit.GetFontNameIndex(cmbFontName.Items.Text);

  //UpdateFontSize
  cmbFontSize.ItemIndex := Edit.GetFontSizeIndex(cmbFontSize.Text, S);
  { this is a litle more complex because cmbFontSize.Items[7] can be changed to
    any size found in i.e. a style }
  if Length(S) > 0
     then begin
        Index := cmbFontSize.ItemIndex;
        cmbFontSize.Items.Text := S;
        cmbFontSize.ItemIndex := Index;
     end;


  //UpdateStylesBox
  cmbStyles.ItemIndex := Edit.GetStylesIndex(cmbStyles.Items.Text);


  //UpdateStatusLine(Edit.ActualElement);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.UpdateButtons;
var
  //dwStatus : Cardinal;
  I: Integer;

  //---------------------------------------
  procedure btnQueryStatus(cmdID: Cardinal; var aButon: TToolButton);
  var
     dwStatus : OLECMDF;
  begin
     dwStatus := Edit.QueryStatus(cmdID);

     aButon.Enabled := (dwStatus and OLECMDF_ENABLED) <> 0;
     aButon.Down := (dwStatus and OLECMDF_LATCHED) <> 0;
   end;
  //---------------------------------------
begin
  //asm int 3 end; //trap
  if PageControl1.ActivePage = tabEdit
     then begin
        btnQueryStatus(IDM_BOLD, btnBold);
        btnQueryStatus(IDM_ITALIC, btnItalic);
        btnQueryStatus(IDM_UNDERLINE, btnUnderline);

        btnQueryStatus(IDM_CUT, btnCut);
        btnQueryStatus(IDM_COPY, btnCopy);
        btnQueryStatus(IDM_PASTE, btnPaste);

        btnQueryStatus(IDM_INDENT, btnIndentRight);
        btnQueryStatus(IDM_OUTDENT, btnIndentLeft);

        btnQueryStatus(IDM_JUSTIFYLEFT, btnAlignLeft);
        btnQueryStatus(IDM_JUSTIFYCENTER, btnAlignCenter);
        btnQueryStatus(IDM_JUSTIFYRIGHT, btnAlignRight);
        btnQueryStatus(IDM_JUSTIFYFULL, btnJustify);

        btnQueryStatus(IDM_BACKCOLOR, btnBGColor);
        btnQueryStatus(IDM_FORECOLOR, btnFGColor);
        btnQueryStatus(IDM_FONT, btnFont);
        if btnFont.enabled and SameText(Edit.ActualElement.Get_tagName, 'FONT')
           then btnFont.Down := true;


        btnQueryStatus(IDM_2D_ELEMENT, btnMakeAbsolute);
        btnQueryStatus(IDM_ORDERLIST, btnNumberedList);
        btnQueryStatus(IDM_UNORDERLIST, btnBulletedList);

        btnQueryStatus(IDM_UNDO, btnUndo);
        btnQueryStatus(IDM_REDO, btnRedo);
        btnQueryStatus(IDM_SHOWZEROBORDERATDESIGNTIME, btnBorders);

        btnQueryStatus(IDM_LocalUndoManager, btnLocalUndo);
        btnQueryStatus(IDM_DROP_UNDO_PACKAGE, btnDropUndo);

        if Coolbar.Bands[2].Visible
           then begin
              //update the table buttons
              if not mnuViewtbl_Table.Enabled
                 then begin
                    btnTableInTable.Enabled := False;
                    btnTableInsRow.Enabled := False;
                    btnTableDelRow.Enabled := False;
                    btnTableInsCol.Enabled := False;
                    btnTableDelCol.Enabled := False;
                    btnTableInsCell.Enabled := False;
                    btnTableDelCell.Enabled := False;
                    btnTableMergeCells.Enabled := False;
                    btnTableTableSplitCells.Enabled := False;
                 end
                 else begin

                    btnQueryStatus(DECMD_INSERTTABLE, btnTableInTable);
                    btnQueryStatus(DECMD_INSERTROW, btnTableInsRow);
                    btnQueryStatus(DECMD_INSERTCOL, btnTableInsCol);
                    btnQueryStatus(DECMD_INSERTCELL, btnTableInsCell);
                    btnQueryStatus(DECMD_DELETEROWS, btnTableDelRow);
                    btnQueryStatus(DECMD_DELETECOLS, btnTableDelCol);
                    btnQueryStatus(DECMD_DELETECELLS, btnTableDelCell);
                    btnQueryStatus(DECMD_MERGECELLS, btnTableMergeCells);
                    btnQueryStatus(DECMD_SPLITCELL, btnTableTableSplitCells);
                 end;
           end;
     end
  else if PageControl1.ActivePage = tabHTML
     then begin
        //mmoHTML
        btnCut.Enabled := mmoHTML.SelLength > 0;
        btnCopy.Enabled := mnuEditCut.Enabled;
        btnPaste.Enabled := true;
        btnUndo.Enabled := mmoHTML.CanUndo;
        //mnuEditRedo.Enabled := false;
     end

  else begin   //preview tab
     btnCopy.Enabled := Preview.Selection;
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnBoldClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_BOLD);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnOpenClick(Sender: TObject);
var
  aFile: string;
begin
  Edit.LoadFile(aFile, true);
  //this command actually returns the loaded file path
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnSaveClick(Sender: TObject);
begin
  Edit.SaveFile;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnInsertImageClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_IMAGE);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnItalicClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_Italic);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnUnderlineClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_Underline);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnFindClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_FIND);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.FormShow(Sender: TObject);
begin

end;
//------------------------------------------------------------------------------
procedure TfrmMain.cmbFontNameChange(Sender: TObject);
begin
  //asm int 3 end; //trap
  if not TEIgnoreChange
     then Edit.CmdSet_S(IDM_FONTNAME, cmbFontName.Text);

  Edit.SetFocusToDoc;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.cmbFontSizeChange(Sender: TObject);
begin
  //asm int 3 end; //trap
  if not TEIgnoreChange
     then Edit.CmdSet_I(IDM_FONTSIZE, cmbFontSize.ItemIndex + 1);

  Edit.SetFocusToDoc;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.cmbStylesChange(Sender: TObject);
begin
  //asm int 3 end; //trap
  if not TEIgnoreChange
     then Edit.SetStyle(cmbStyles.Text);

  Edit.SetFocusToDoc;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.cmbDropDown(Sender: TObject);
begin
  FFocusObject := Sender as TComboBox;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.cmbFontSizeDropDown(Sender: TObject);
var
  I: Integer;
  S: String;
begin
  //dont show anything than font size 1 to 7 in the drop down box
  For I := cmbFontSize.Items.Count -1  downto 0 do
     begin
        //ALL std. font size looks like theis: '2  (10 pt)'
        S := copy(cmbFontSize.Items[I], 1, 4);
        if (Length(S) < 4) or
           (pos(S[1], '1234567') = 0) or
           (Copy(S, 2, 3) <> '  (')
           then cmbFontSize.Items.Delete(I)
     end;

  FFocusObject := Sender as TComboBox;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.tlb_FormatMouseMove(Sender: TObject; Shift: TShiftState; X, Y: Integer);
begin
  if (not assigned(FFocusObject)) or
     (FFocusObject.DroppedDown)
     then exit;

  //now the box isen dropped down
  //if the mouse leavs the object, then shift focus back to Edit

  if (Y < FFocusObject.Left) or
     (X > FFocusObject.Top) or
     (Y > FFocusObject.Left + FFocusObject.Width) or
     (X < FFocusObject.Top - FFocusObject.Height)
     then begin
        Edit.SetFocusToDoc;
        FFocusObject := nil;
     end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuFileExitClick(Sender: TObject);
begin
  Close;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuQueryStatus(cmdID: Cardinal; var amnu: TMenuItem);
var
  dwStatus : OLECMDF;
begin
  dwStatus := Edit.QueryStatus(cmdID);
  amnu.Enabled := (dwStatus and OLECMDF_ENABLED) <> 0;
  amnu.Checked := (dwStatus and OLECMDF_LATCHED) <> 0;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableClick(Sender: TObject);
          //prepare Table menu before it drops down
begin
  mnuQueryStatus(DECMD_INSERTTABLE, mnuTableInsertTable);
  mnuQueryStatus(DECMD_INSERTROW, mnuTableInsertRow);
  mnuQueryStatus(DECMD_INSERTCOL, mnuTableInsertCol);
  mnuQueryStatus(DECMD_INSERTCELL, mnuTableInsetCell);
  mnuQueryStatus(DECMD_DELETEROWS, mnuTableDeleteRow);
  mnuQueryStatus(DECMD_DELETECOLS, mnuTableDeleteCol);
  mnuQueryStatus(DECMD_DELETECELLS, mnuTableDeleteCell);
  mnuQueryStatus(DECMD_MERGECELLS, mnuTableMergeCells);
  mnuQueryStatus(DECMD_SPLITCELL, mnuTableSplitCells);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnInsertTableClick(Sender: TObject);
var
  ov: OleVariant;
begin
  //asm int 3 end; //trap

  frmTable := TfrmTable.Create(self);
  try
     if frmTable.ShowModal = mrOK
        then begin
           { IDM_INSERTTABLE is called with a 7 param safe array
             0 =  Number of rows
             1 =  Number of columns
             2 =  Table width in %
             3 =  Table tag attributes
             4 =  Cell tag attributes
             5 =  Table caption
             6 =  Wrap text in cells }

           ov := VarArrayCreate([0, 6], varVariant);
           ov[0] := StrToInt(frmTable.edRows.Text);
           ov[1] := StrToInt(frmTable.edCols.Text);
           ov[2] := frmTable.edWidth.Text + '%';
           ov[3] := 'Border=' + frmTable.edTBorder.text +       // Table Attributes
                    ' cellPadding=' + frmTable.edPad.text +
                    ' cellSpacing=' + frmTable.edSpace.Text;
           ov[4] := 'bgColor=' + frmTable.cmbTableBGColor.text; // Cell Attributes
           ov[5] := frmTable.edCaption.text;
           ov[6] := frmTable.chkWrap.checked;

           frmMain.Edit.CmdSet(DECMD_INSERTTABLE, ov);
        end;
  finally
     frmTable.Free;
  end;

(* Old code
//insert a table in the DHTML Editor
var
  insertTableParam: DEInsertTableParam;
  ovInsertTableParam: OleVariant;
  YNWrap: string;
begin
  if frmTable.ShowModal = mrOK then
  begin
    if frmTable.chkWrap.checked then
      YNWrap := ' noWrap'
    else
      YNWrap := ' Wrap';

    insertTableParam := CreateComObject(Class_DEInsertTableParam) as IDEInsertTableParam;

    insertTableParam.NumRows := StrToInt(frmTable.edRows.Text);
    insertTableParam.NumCols := StrToInt(frmTable.edCols.Text);
    insertTableParam.TableAttrs := ' width=' + frmTable.edWidth.Text + '%' +
      ' Border=' + frmTable.edTBorder.text +
      ' cellPadding=' + frmTable.edPad.text +
      ' cellSpacing=' + frmTable.edSpace.Text;
    insertTableParam.CellAttrs := ' bgColor=' + frmTable.cmbTableBGColor.text + YNWrap;
    insertTableParam.Caption := frmTable.edCaption.text;

    ovInsertTableParam := OleVariant(insertTableParam);
    ExecCommand(DECMD_INSERTTABLE, OLECMDEXECOPT_DODEFAULT, ovInsertTableParam);

    mnuViewtbl_Table.Enabled:= True;
  end;
  *)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditDeleteClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_DELETE)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditCutClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_CUT)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditCopyClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_COPY);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditPasteClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_PASTE);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuInsertHyperlinkClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_HYPERLINK);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnInsertLinkClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_HYPERLINK);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnAlignCenterClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_JUSTIFYCENTER)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnAlignLeftClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_JUSTIFYLEFT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnAlignRightClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_JUSTIFYRIGHT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnJustifyClick(Sender: TObject);
begin
  if (Edit.QueryStatus(IDM_JUSTIFYFULL) and OLECMDF_LATCHED) <> 0
     then Edit.CmdSet(IDM_JUSTIFYNONE)
     else Edit.CmdSet(IDM_JUSTIFYFULL);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditClick(Sender: TObject);
//prepare Edit menu before it drops down
var
  I: Integer;
begin
  //asm int 3 end; //trap

  for I := 0 to mnuEdit.Count -1 do
     mnuEdit.Items[I].Enabled := false;

  if PageControl1.ActivePage = tabEdit
     then begin
        mnuEditUndo.Enabled := btnUndo.Enabled;
        mnuEditRedo.Enabled := btnRedo.Enabled;
        mnuEditCut.Enabled := Edit.Selection;
        mnuEditCopy.Enabled := Edit.Selection;
        mnuEditPaste.Enabled := Edit.Selection;
        mnuEditDelete.Enabled := true;
        mnuEditSelectAll.Enabled := true;
        mnuEditFind.Enabled := true;
        mnuEditUnlink.Enabled := Edit.QueryEnabled(IDM_UNLINK);
     end

  else if PageControl1.ActivePage = tabHTML
     then begin
        //mmoHTML
        mnuEditUndo.Enabled := mmoHTML.CanUndo;
        mnuEditCut.Enabled := mmoHTML.SelLength > 0;
        mnuEditCopy.Enabled := mnuEditCut.Enabled;
        mnuEditPaste.Enabled := true;
        mnuEditDelete.Enabled := true;
        mnuEditSelectAll.Enabled := true;
     end

  else begin   //preview tab
     mnuEditCopy.Enabled := Preview.Selection;
     mnuEditSelectAll.Enabled := true;
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditRedoClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_REDO)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditUndoClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_UNDO)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditSelectAllClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_SELECTALL);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditUnlinkClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_UNLINK)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableInsertRowClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_INSERTROW);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableInsertColClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_INSERTCOL);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableInsetCellClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_INSERTCELL);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableDeleteRowClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_DELETEROWS);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableDeleteColClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_DELETECOLS);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableDeleteCellClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_DELETECELLS);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnFGColorClick(Sender: TObject);
var
  Ov: OleVariant;
begin
  //asm int 3 end; //trap

  if ColorDialog.Execute then
  begin
    Ov := Color2HTML(ColorDialog.Color);
    Edit.CmdSet(IDM_FORECOLOR, Ov);;
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableMergeCellsClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_MERGECELLS);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuTableSplitCellsClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_SPLITCELL);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnIndentRightClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_INDENT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnIndentLeftClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_OUTDENT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnBulletedListClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_UNORDERLIST);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnNumberedListClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_ORDERLIST);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionAboveTextClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_BRING_ABOVE_TEXT); // Set Z-INDEX to 100
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionBringForwardClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_BRING_FORWARD); //increment Z-INDEX by one, skip from -99 to 99
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionBringtoFrontClick(Sender: TObject);
begin
   // Set Z-INDEX to 1 higher that any other Z-INDEX on page, min 100
  Edit.CmdSet(DECMD_BRING_TO_FRONT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionBelowTextClick(Sender: TObject);
begin
  //SendbehindHTMLstream
  Edit.CmdSet(DECMD_SEND_BELOW_TEXT); // Set Z-INDEX to -100
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionSendtoBackClick(Sender: TObject);
begin
  // Set Z-INDEX to 1 lower that any other Z-INDEX on page, min -100
  Edit.CmdSet(DECMD_SEND_TO_BACK);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionSendBackwardClick(Sender: TObject);
begin
  //decrement Z-INDEX by one, skip from 99 to -99
  Edit.CmdSet(DECMD_SEND_BACKWARD);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuLockElementClick(Sender: TObject);
begin
  Edit.CmdSet(DECMD_LOCK_ELEMENT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuConstrainElementClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_CONSTRAIN);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnBGColorClick(Sender: TObject);
var
  Ov: OleVariant;
begin
  //asm int 3 end; //trap
  if ColorDialog.Execute then
  begin
    Ov := Color2HTML(ColorDialog.Color);
    Edit.CmdSet(IDM_BACKCOLOR, Ov);
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnFontClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_FONT);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionClick(Sender: TObject);
begin
  //asm int 3 end; //trap
  mnuPositionMakeAbsolute.Enabled := btnMakeAbsolute.Enabled;
  mnuPositionMakeAbsolute.checked := btnMakeAbsolute.Down;

  if mnuPositionMakeAbsolute.checked
     then mnuQueryStatus(DECMD_LOCK_ELEMENT, mnuLockElement)
     else begin
        mnuLockElement.Enabled := false;
        mnuLockElement.Checked := false;
     end;

  mnuQueryStatus(DECMD_BRING_ABOVE_TEXT, mnuPositionAboveText);
  mnuQueryStatus(DECMD_BRING_FORWARD, mnuPositionBringForward);
  mnuQueryStatus(DECMD_BRING_TO_FRONT, mnuPositionBringToFront);
  mnuQueryStatus(DECMD_SEND_BELOW_TEXT, mnuPositionBelowText);
  mnuQueryStatus(DECMD_SEND_TO_BACK, mnuPositionSendToBack);
  mnuQueryStatus(DECMD_SEND_BACKWARD, mnuPositionSendBackward);

  mnuQueryStatus(IDM_CONSTRAIN, mnuConstrainElement);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionMakeAbsoluteClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_ABSOLUTE_POSITION);
end;
//------------------------------------------------------------------------------
{W3C: The following properties have been deprecated in favor of the corresponding
ones for the BODY element: alinkColor, bgColor, fgColor, linkColor, vlinkColor
and background-image}
procedure TfrmMain.mnuFilePropertiesClick(Sender: TObject);
var
  BGImageFile: string;
  DOC: IHTMLDocument2;
  Ov: OleVariant;
begin
  frmPageProp := TfrmPageProp.Create(Self);
  try
     DOC := Edit.DOM;
     frmPageProp.edTitle.Text := DOC.Title;
     frmPageProp.cmbBackColor.Text := DOC.bgColor;
     frmPageProp.cmbALinkColor.Text := DOC.alinkcolor;
     frmPageProp.cmbVLinkColor.Text := DOC.vlinkcolor;
     frmPageProp.cmbLinkColor.Text := DOC.linkcolor;
     BGImageFile := (DOC.body as IHTMLBodyElement).background;

     if frmPageProp.showModal = mrOK
        then begin
           DOC.Title := frmPageProp.edTitle.Text;
           Caption:=Application.Title + ' [' + DOC.Title + ']';

           Ov := 'darkorange';
           DOC.Set_bgColor(Ov);// frmPageProp.cmbBackColor.Text;
           DOC.alinkcolor:=frmPageProp.cmbALinkColor.Text;
           DOC.vlinkcolor:=frmPageProp.cmbVLinkColor.Text;
           DOC.linkcolor:=frmPageProp.cmbLinkColor.Text;
           BGImageFile := frmPageProp.OpenPictureDialog.Filename;
           //if no image is assigned, no bgimage-tag is placed in the body tag
           if BGImageFile <> ''
              then (DOC.body as IHTMLBodyElement).background := BGImageFile;
        end;
  finally
     frmPageProp.Free;
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnPrintClick(Sender: TObject);
begin
  //asm int 3 end; //trap
  Print(False);  //show no dialog
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPrintpreviewClick(Sender: TObject);
begin
  //asm int 3 end; //trap

  if not mnuPrintpreview.enabled
     then exit;

  if PageControl1.ActivePage = tabPreview
     then Preview.PrintPreview(EmptyParam)

  else if PageControl1.ActivePage = tabEdit
     then Edit.PrintPreview(EmptyParam);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.PrintHTMLSource;
var
  Buffer: PChar;
  Size: Integer;
  Prn: PrnRec;
begin
  if mmoHTML.SelLength > 0
     then begin
        PrintDialog1.Options := [poSelection];
        PrintDialog1.PrintRange := prSelection;
     end
     else PrintDialog1.Options := [];

  if PrintDialog1.Execute
     then begin // determine the range the user wants to print
        prn.headerText := 'Just Testing';
        prn.headerFontName := 'Times New Roman';
        prn.headerFontSize := 14;
        prn.headerFontStyle := [fsUnderline, fsBold];
        prn.pageNumber := 1;
        Prn.bodyFontName := 'Arial';
        Prn.bodyFontSize := 10;
        Prn.bodyFontStyle := [];
        Prn.footerFontName := 'Times New Roman';
        Prn.footerFontSize := 8;
        Prn.footerFontStyle := [fsItalic];
        Prn.inchesInLeftMargin := 1.0;
        Prn.inchesInRightMargin := 1.0;
        Prn.inchesInTopMargin := 0.5;
        Prn.inchesInBottomMargin := 0.5;
        Prn.lineSpacing := 1.5;
        Prn.paperWidth := 8.5;
        Prn.paperHeight := 11.0;

        Printer.BeginDoc;
        try
           if PrintDialog1.PrintRange = prAllPages
              then PrintString(Prn, PChar(mmoHTML.Lines.Text))

           else if PrintDialog1.PrintRange = prSelection
              then begin
                 Size := mmoHTML.SelLength;
                 Inc(Size);
                 GetMem(Buffer, Size);
                 try
                    mmoHTML.GetSelTextBuf(Buffer, Size);
                    PrintString(Prn, Buffer);
                 finally
                    FreeMem(Buffer, Size);
                 end;
              end;
         finally
           Printer.EndDoc;
         end;
     end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.Print(ShowDlg: boolean = True);
var
  Ov: OleVariant;
begin
  //asm int 3 end; //trap
  Ov := ShowDlg;

  if PageControl1.ActivePage = tabHTML
     then PrintHTMLSource

  else if PageControl1.ActivePage = tabPreview  //print from preview
     then Preview.PrintDocument(Ov)

  else Edit.PrintDocument(Ov);
end;
//------------------------------------------------------------------------------

//------------------------------------------------------------------------------
procedure TfrmMain.mnuHelpAboutClick(Sender: TObject);
begin
  ShowMessage('MSHTML Editor' + #13#10 +
    'Version 3.6 for MS IE 5.5 or higher' + #13#10 +
    'Copyright (C) Kurt Senfer'+ #13#10 +
    'email: Kurt.Senfer@siemens.com' + #13#10 + #13#10);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuInsertHTMLClick(Sender: TObject);
var
  ovSelection: OleVariant;
  ovTextRange: OleVariant;
begin

  frmInsertHTML := TfrmInsertHTML.Create(Self);
  try
     if frmInsertHTML.ShowModal = mrOK
        then begin
           //this demonstrates that KsDHTMLEDLib supports Dispatch calls

           // get the IE selection object
           ovSelection := Edit.OleObject.Document.selection;

           // create a TextRange from the current selection
           ovTextRange := ovSelection.createRange;

           // paste our html into the range
           ovTextRange.pasteHTML(frmInsertHTML.InsHTML.Lines.Text);
     end;
  finally
     frmInsertHTML.free;
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnNewClick(Sender: TObject);
begin
  Edit.NewDocument;

  If Coolbar.Bands[2].Visible
     then Coolbar.Bands[2].Visible:= False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuPositionGridSettingsClick(Sender: TObject);
begin
  DoDlg_Grid;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuFileSaveAsClick(Sender: TObject);
var
  aFile: String;
begin
  //asm int 3 end; //trap
  Edit.SaveFileAs(aFile);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.PageControl1Changing(Sender: TObject; var AllowChange: Boolean);
begin
  //asm int 3 end; //trap

  if (PageControl1.ActivePage = tabHTML) and //changing from Source HTML
      mmoHTML.Modified
     then Edit.DocumentHTML := mmoHTML.Text

  else if PageControl1.ActivePage = tabPreview  //Changing from preview
     then PreView.AssignDocument //Clear old contenet from preview browser

  else tlb_Format.Visible := False; //changing from WYSIWYG
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnPrevInBrowserClick(Sender: TObject);
var//14/05/99 RuudK - Not perfect yet, saveas called twice...
  Directory : string;
  ovFileName: OleVariant;
begin
  ovFileName := _CurrentDocumentPath; //DHTMLEdit.CurrentDocumentPath;
  if ovFileName = '' then
  begin
    if MessageDlg('The document has not been saved yet. Save now?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      MessageDlg('Saving the document to disk.', mtInformation,
        [mbOk], 0);
      btnSaveClick(Sender);
      FmxUtils.ExecuteFile(ovFileName, '', Directory, SW_SHOW);
    end;
    if MessageDlg('The document has not been saved yet. Save now?',
      mtConfirmation, [mbYes, mbNo], 0) = mrNo then Exit;
  end;
  if Changed = True then
  begin
    if MessageDlg('The content has been changed. Save now?',
      mtConfirmation, [mbYes, mbNo], 0) = mrYes then
    begin
      MessageDlg('Saving the document to disk.', mtInformation,
        [mbOk], 0);
      btnSaveClick(Sender);
      FmxUtils.ExecuteFile(ovFileName, '', Directory, SW_SHOW);
    end;
    if MessageDlg('The content has been changed. Save now?',
      mtConfirmation, [mbYes, mbNo], 0) = mrNo then Exit;
  end;
  if Changed = False then
  begin
    FmxUtils.ExecuteFile(ovFileName, '', Directory, SW_SHOW);
  end;
  mmoHTML.Modified := False;
  Changed := False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mmoHTMLChange(Sender: TObject);
begin
  Changed:=True;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.FormCreate(Sender: TObject);
begin
  Caption := Application.Title + ' [No Title]';
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuViewtbl_TableClick(Sender: TObject);
begin
  mnuViewtbl_Table.Checked:= not mnuViewtbl_Table.Checked;
  if mnuViewtbl_Table.Checked
     then Coolbar.Bands[2].Visible:= True;

  if mnuViewtbl_Table.Checked = False
     then Coolbar.Bands[2].Visible:= False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuViewtbl_FormatClick(Sender: TObject);
begin
  mnuViewtbl_Format.Checked:= not mnuViewtbl_Format.Checked;
  if mnuViewtbl_Format.Checked
     then Coolbar.Bands[1].Visible:= True;

  if mnuViewtbl_Format.Checked = False
     then Coolbar.Bands[1].Visible:= False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuViewtbl_EditClick(Sender: TObject);
begin
  mnuViewtbl_Edit.Checked:= not mnuViewtbl_Edit.Checked;
  if mnuViewtbl_Edit.Checked
     then Coolbar.Bands[0].Visible:= True;

  if mnuViewtbl_Edit.Checked = False
     then Coolbar.Bands[0].Visible:= False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.LookForObjects;
begin
  //asm int 3 end; //trap
  //tags('TABLE'); //collection of all tables - if anny

  if (Edit.DOM <> nil) and
     ((Edit.DOM.all.tags('TABLE') as IHTMLElementCollection).length > 0)
     then mnuViewtbl_Table.Enabled:= True
     else mnuViewtbl_Table.Enabled:= False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.mnuEditRemoveTagsClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_RemoveFormat)
end;
//------------------------------------------------------------------------------
procedure TfrmMain.PageControl1Change(Sender: TObject);
var
  SelStart: Integer;
  SelEnd: Integer;
  CurLine: Integer;
  CurTop: Integer;
  S: String;

   //---------------------------------------
   procedure EnableAllButtons(Enable: Boolean);
   var
     I: Integer;
   begin
     for I := 0 to tlb_Edit.ButtonCount -1 do
        tlb_Edit.Buttons[I].Enabled := Enable;
   end;
   //---------------------------------------
begin
  //asm int 3 end; //trap


  //change to WYSIWYG
  if PageControl1.ActivePage = tabEdit
     then begin
        EnableAllButtons(true); //Enable all buttons
        mnuPrintpreview.enabled := true;
        //DHTMLEdit.ShowSpcGlyphs;

        tlb_Format.Visible := True;
        //LastPage := tabEdit;
        if SourceView
           then Edit.SyncDOC(mmoHTML.text, mmoHTML.SelStart, mmoHTML.SelStart + mmoHTML.SelLength);

        SourceView := false;

        Edit.SetFocusToDoc;
        exit;
     end;

  EnableAllButtons(false); //UnEnable all buttons

  //change view to Source HTML
  if PageControl1.ActivePage = tabHTML
     then begin
        mnuPrintpreview.enabled := false;
        { put the HTML source in the memo and place the cursor or selection
          at the same pos in memo as in the WYSIWYG }

        mmoHTML.Text := Edit.SelectedDocumentHTML(SelStart, SelEnd);
        mmoHTML.Selstart := SelStart;
        mmoHTML.SelLength := SelEnd - SelStart;
        mmoHTML.SetFocus; //or the heighleighted text might not show

        //Scroll line with cursor to mid scren
        CurLine := SendMessage(mmoHTML.Handle, EM_LINEFROMCHAR, mmoHTML.SelStart, 0);
        CurTop := SendMessage(mmoHTML.Handle, EM_GETFIRSTVISIBLELINE, 0, 0);
        SendMessage(mmoHTML.Handle, EM_LINESCROLL, 0, -((CurTop - CurLine) div 2));

        SourceView := true;
        exit;
     end;

  //Change to preview
  if (PageControl1.ActivePage = tabPreview)
     then begin
        mnuPrintpreview.enabled := true;
        S := Edit.BaseURL;
        if Length(S) > 0
           then Preview.BaseURL := S
           else Preview.BaseURL := Edit.CurrentDocumentPath;
           
        Preview.DocumentHTML := Edit.DocumentHTML;
     end;
end;
//-----------------------------------------------------------------------------
procedure TfrmMain.EditDocumentComplete(Sender: TObject);
var
  FItems: TStringList;
  I: Integer;
begin
  if not Edit.BrowseMode
     then begin
         FrmMain.cmbStyles.Items.Clear;
         TEIgnoreChange := True;

         //Get standard styles
         cmbStyles.Items.Text := Edit.GetBuildInStyles;

         //get attached stylesheets styles
         FItems := TStringList.create;
         try
           FItems.Text := Edit.GetExternalStyles;
           for I := 1 to FItems.Count -1 do
              cmbStyles.Items.Add(FItems[I]);
         finally
           FItems.Free;
         end;

         TEIgnoreChange := False;
  end;

  If Edit.DocumentTitle = ''
     then Edit.DocumentTitle := 'No Title';

  Caption := Application.Title + ' [' + Edit.DocumentTitle + ']';
end;
//-----------------------------------------------------------------------------
procedure TfrmMain.EditShowContextMenuEx(Sender: TObject; xPos, yPos: Integer);
begin
  { we could manage the popup menu heir, but as we consume the ONCONTEXTMENU
    event in the EditDesigners PreHandleEvent we never get heir anymore }

  PopUpMenu1.Popup(xPos, yPos);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnShowDetailsClick(Sender: TObject);
begin
  Edit.ShowDetails := not Edit.ShowDetails;
  btnShowDetails.Down := Edit.ShowDetails;
  if Edit.ShowDetails
     then btnShowDetails.Hint := 'Hide glyphs'
     else btnShowDetails.Hint := 'Show glyphs';
end;
//------------------------------------------------------------------------------
function TfrmMain.EditQueryService(const rsid, iid: TGUID; out Obj: IUnknown): HRESULT;
begin
  { skeleton - just in case we want to experiment }

  Obj := nil;
  result := E_NOINTERFACE;
end;
//------------------------------------------------------------------------------
var
  Borders: boolean = false;
//------------------------------------------------------------------------------
procedure TfrmMain.btnBordersClick(Sender: TObject);
var
  Ov: OleVariant;
begin
  Borders := not Borders;
  btnBorders.Down := Borders;
  ov := Borders;
  Edit.CmdSet(IDM_SHOWZEROBORDERATDESIGNTIME, ov);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnLocalUndoClick(Sender: TObject);
begin
  Edit.CmdSet_B(IDM_LocalUndoManager, not Edit.LocalUndoManager);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.btnDropUndoClick(Sender: TObject);
begin
  Edit.CmdSet(IDM_DROP_UNDO_PACKAGE);
end;
//------------------------------------------------------------------------------
procedure TfrmMain.EditPostEditorEventNotify(Sender: TObject; inEvtDispId: Integer; const pIEventObj: IHTMLEventObj; var aResult: HRESULT);
begin
  //every event comes heir after its done
  aResult := S_FALSE;

  case inEvtDispId of
     DISPID_HTMLELEMENTEVENTS_ONMOUSEMOVE: exit; {no need to call DHTMLEditDisplayChanged
                                                  on each mouse move }
  end;

  Statusbar1.Panels[0].Text:= (pIEventObj.Get_srcElement).OuterHTML;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.EditPreHandleEvent(Sender: TObject; inEvtDispId: Integer; const pIEventObj: IHTMLEventObj; var aResult: HRESULT);
var
  P: tagPOINT;
  aDisplays: IDisplayServices;
begin
  aResult := S_FALSE;

  //every event comes heir before it is hanled - return S_OK consumes the event
  case inEvtDispId of
     DISPID_HTMLELEMENTEVENTS_ONCONTEXTMENU:
        begin
           aresult := S_OK;

           //get click position
           P.X :=  pIEventObj.ScreenX;
           P.Y :=  pIEventObj.ScreenY;

           //transform click position to a screen pos
           aDisplays := Edit.DOM as IDisplayServices;
           aDisplays.TransformPoint(P, COORD_SYSTEM_CONTAINER, COORD_SYSTEM_GLOBAL, nil);

           PopUpMenu1.Popup(P.X, P.Y);
        end;
  end;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.EditInitialize(Sender: TObject; var NewFile: String);
begin
  { return the name of a file you want to open at startup
    i.e.: NewFile := Paramstr(1);
    or:   NewFile := 'V:\KHDev\KnowHow\000-304.htm';

    do other initialization }

  Edit.SetDebug(true);
  PreView.SetDebug(true);
  NewFile := Paramstr(1);

  TEIgnoreChange := True;
  cmbFontName.Items.text := Edit.GetFonts;  
  TEIgnoreChange := False;
end;
//------------------------------------------------------------------------------
procedure TfrmMain.FormClose(Sender: TObject; var Action: TCloseAction);
begin
  if Edit.EndCurrentDoc(true) <> S_OK  //EndCurrentDoc({Cancel is posible})
     then Action := caNone;
end;
//------------------------------------------------------------------------------


procedure TfrmMain.SpeedButton1Click(Sender: TObject);
{//insert a table in the DHTML Editor
var
  insertTableParam: DEInsertTableParam;
  ovInsertTableParam: OleVariant;
  YNWrap: string; }
begin
{  if frmTable.ShowModal = mrOK then
  begin
    if frmTable.chkWrap.checked then
      YNWrap := ' noWrap'
    else
      YNWrap := ' Wrap';

    insertTableParam := CreateComObject(Class_DEInsertTableParam) as IDEInsertTableParam;

    insertTableParam.NumRows := StrToInt(frmTable.edRows.Text);
    insertTableParam.NumCols := StrToInt(frmTable.edCols.Text);
    insertTableParam.TableAttrs := ' width=' + frmTable.edWidth.Text + '%' +
      ' Border=' + frmTable.edTBorder.text +
      ' cellPadding=' + frmTable.edPad.text +
      ' cellSpacing=' + frmTable.edSpace.Text;
    insertTableParam.CellAttrs := ' bgColor=' + frmTable.cmbTableBGColor.text + YNWrap;
    insertTableParam.Caption := frmTable.edCaption.text;

    ovInsertTableParam := OleVariant(insertTableParam);
    ExecCommand(DECMD_INSERTTABLE, OLECMDEXECOPT_DODEFAULT, ovInsertTableParam);

    mnuViewtbl_Table.Enabled:= True;
  end;  }
end;

procedure TfrmMain.SpeedButton2Click(Sender: TObject);
var
  InsertTableParam: DEInsertTableParam;
  ovInsertTableparam: OleVariant;
begin
  // Edit.DOM
  // Carregar/Passar texto simples : Edit.DOC.body.innerText := Memo1.Lines.Text;

  // Passar Código HTML do body: Memo1.Lines.Text := Edit.DOC.body.innerHTML;
  // Passa o código HTML: Memo1.Lines.Text := Edit.DocumentHTML;

//  Memo1.Lines.Text := Edit.DocumentHTML;

//  Edit.DocumentHTML := '<html>Hello</html>';


//   Memo1.Lines.Text := Edit.DOC.body.innerText;
  Edit.DOC.execCommand('InsertRow', true, xxx);
end;

end.

