{ "Copyright (C) 2002-2004  Kurt Senfer <Kurt.Senfer@siemens.com>"

  The code in this unit is released under:
  GNU General Public License - http://www.gnu.org/copyleft/gpl.html }

unit PrintText;

interface

uses Windows, graphics;

type
  PrnRec = record
     Cur: TPoint;            //Current print pos
     Finish: TPoint;         //End of the printable area
     pageNumber: Integer;
     headerText: String;
     headerFontName: String;
     headerFontSize: Integer;
     headerFontStyle: TFontStyles;
     bodyFontName: String;
     bodyFontSize: Integer;
     bodyFontStyle: TFontStyles;
     footerFontName: String;
     footerFontSize: Integer;
     footerFontStyle: TFontStyles;
     inchesInLeftMargin: Real;
     inchesInRightMargin: Real;
     inchesInTopMargin: Real;
     inchesInBottomMargin: Real;
     lineSpacing: Real;
     paperWidth: Real;
     paperHeight: Real;
     LeftMargin: integer;
     yHeader: integer;
     yBody: integer;
     yFooter: integer;
     lineHeight: integer;
  end;


procedure PrintString(var Prn: PrnRec; Text: PChar);


implementation

uses Messages, SysUtils, Printers, KS_Procs;



//----------------------------------------------------------------
function CharHeight: Word;
var
  Metrics: TTextMetric;
begin
  GetTextMetrics(Printer.Canvas.Handle, Metrics);
  Result := Metrics.tmHeight;
end;
//----------------------------------------------------------------
procedure PrintHeader(var Prn: PrnRec);

begin
  Printer.Canvas.Font.Name := Prn.headerFontName;
  Printer.Canvas.Font.Size := Prn.headerFontSize;
  Printer.Canvas.Font.Style := Prn.headerFontStyle;
  Printer.Canvas.TextOut(Prn.LeftMargin, Prn.yHeader, prn.headerText);
  Prn.Cur.X := Prn.LeftMargin;
  Prn.Cur.Y := Prn.yBody;

  //set up body
  Printer.Canvas.Font.Name := Prn.bodyFontName;
  Printer.Canvas.Font.Size := Prn.bodyFontSize;
  Printer.Canvas.Font.Style := Prn.bodyFontStyle;

  if Prn.lineSpacing = 0
     then Prn.LineHeight := CharHeight //default line height
     else Prn.LineHeight := Round(CharHeight * Prn.lineSpacing);
end;
//----------------------------------------------------------------
procedure PrintFooter(var Prn: PrnRec);
begin
  Printer.Canvas.Font.Name := Prn.footerFontName;
  Printer.Canvas.Font.Size := Prn.footerFontSize;
  Printer.Canvas.Font.Style := Prn.footerFontStyle;
  Printer.Canvas.TextOut(Prn.LeftMargin, Prn.yFooter, 'Printed by <your name>   '
                                               + IntToStr(prn.pageNumber));
  Inc(Prn.PageNumber);
end;
//----------------------------------------------------------------
procedure PageSetup (var prn: PrnRec);
var
  lineHeight: integer;
  RightMargin: integer;
  presetTopBottomMargin, presetLeftRightMargin: real;
begin
  //Compute height of the header line
  Printer.Canvas.Font.Name := Prn.headerFontName;
  Printer.Canvas.Font.Size := Prn.headerFontSize;
  Printer.Canvas.Font.Style := Prn.headerFontStyle;

  lineHeight := round(CharHeight * Prn.lineSpacing);

  {Compute yHeader, subtracting the margin already provided,
   and yBody, using lineHeight of the header}
  presetTopBottomMargin := (Printer.Canvas.Font.PixelsPerInch * Prn.paperHeight - Printer.PageHeight)/2;
  Prn.yHeader := round(Printer.Canvas.Font.PixelsPerInch * Prn.inchesInTopMargin - presetTopBottomMargin);
  Prn.yBody := Prn.yHeader + 2 * lineHeight;

  {Compute xLeftMargin, subtracting the margin already provided.}
  presetLeftRightMargin := (Printer.Canvas.Font.PixelsPerInch * Prn.paperWidth - Printer.PageWidth)/2;
  Prn.LeftMargin := round(Printer.Canvas.Font.PixelsPerInch * Prn.inchesInLeftMargin - presetLeftRightMargin);
  RightMargin := round(Printer.Canvas.Font.PixelsPerInch * Prn.inchesInRightMargin - presetLeftRightMargin);

  {Compute yFooter based on footer font.}
  Printer.Canvas.Font.Name := Prn.footerFontName;
  Printer.Canvas.Font.Size := Prn.footerFontSize;
  Printer.Canvas.Font.Style := Prn.footerFontStyle;

  lineHeight := round(CharHeight * Prn.lineSpacing);
  presetTopBottomMargin := (Printer.Canvas.Font.PixelsPerInch * Prn.paperHeight - Printer.PageHeight)/2;
  Prn.yFooter := Printer.PageHeight - round(Printer.Canvas.Font.PixelsPerInch * Prn.inchesInBottomMargin - presetTopBottomMargin) - lineHeight;

  Prn.Cur.X := Prn.LeftMargin;
  Prn.Cur.Y := Prn.yBody;
  Prn.Finish.X := Printer.PageWidth - RightMargin;
  Prn.Finish.Y := Prn.yFooter;
end;
//----------------------------------------------------------------
procedure NewPage(var Prn: PrnRec);
begin
  PrintFooter(Prn);
  Printer.NewPage;
  PrintHeader(Prn);
end;
//----------------------------------------------------------------
procedure NewLine(var Prn: PrnRec);
begin
  Prn.Cur.X := Prn.LeftMargin;    //X start pos of new line
  Inc(Prn.Cur.Y, Prn.LineHeight); //Y start pos of new line

  //if no more lines left start a new page
  if Prn.Cur.Y > (Prn.Finish.Y - (Prn.LineHeight * 2))
     then NewPage(Prn);
end;
//----------------------------------------------------------------
procedure PrintLine(var Prn: PrnRec; Text: PChar; Len: Integer);
var
  Extent: TSize;
  L, L2: Integer;
  S: String;
begin
  while Len > 0 do  //print it all
     begin
        L := Len;

        //do we need to brak line ?
        GetTextExtentPoint(Printer.Canvas.Handle, Text, L, Extent);

        while (L > 0) and (Extent.cX + Prn.Cur.X > Prn.Finish.X) do
           begin
              L := CharPrev(Text, Text + L) - Text;
              GetTextExtentPoint(Printer.Canvas.Handle, Text, L, Extent);
           end;

        if L < Len
           then begin  //try to find a better place to break text
              S := Copy(Text, L - 15, 15);
              L2 := intMax(Pos(' ', S), Pos('>', S));
              if L2 > 0
                 then begin
                    dec(L, 15 - L2 + 1);
                    GetTextExtentPoint(Printer.Canvas.Handle, Text, L, Extent);
                 end;
           end;


        if Extent.cY > CharHeight  //unexpected heigh character in line
           then Prn.Cur.Y := Prn.Cur.Y + (Extent.cY - CharHeight) + 2;

        Windows.TextOut(Printer.Canvas.Handle, Prn.Cur.X, Prn.Cur.Y, Text, L);

        Dec(Len, L);
        Inc(Text, L);
        if Len > 0           //if more text then break line
           then NewLine(Prn)
           else Inc(Prn.Cur.X, Extent.cX);
     end;
end;
//----------------------------------------------------------------
procedure PrintString(var Prn: PrnRec; Text: PChar);
var
  L: Integer;
  TabWidth: Word;
  Len: Integer;

  //-----------------------------
  procedure Flush;
  begin
     if L <> 0
        then PrintLine(Prn, Text, L);
     Inc(Text, L + 1);
     Dec(Len, L + 1);
     L := 0;
  end;
  //-----------------------------
  function AvgCharWidth: Word;
  var
     Metrics: TTextMetric;
  begin
     GetTextMetrics(Printer.Canvas.Handle, Metrics);
     Result := Metrics.tmAveCharWidth;
  end;
  //-----------------------------
begin
  PageSetup(Prn);
  PrintHeader(Prn);
  Len := length(Text);
  L := 0;

  while L < Len do
     begin
        case Text[L] of
           #9:
              begin
                 Flush;
                 TabWidth := AvgCharWidth * 8;
                 Inc(Prn.Cur.X, TabWidth - ((Prn.Cur.X + TabWidth + 1) mod TabWidth) + 1);
                 if Prn.Cur.X > Prn.Finish.X
                    then NewLine(Prn);
              end;
           #13: Flush;
           #10:
              begin
                 Flush;
                 NewLine(Prn);
              end;
           ^L:
              begin
                 Flush;
                 NewPage(Prn);
              end;
           else Inc(L);
        end; //case
     end; //while

  Flush;
  PrintFooter(Prn);
end;
//----------------------------------------------------------------

end.
 