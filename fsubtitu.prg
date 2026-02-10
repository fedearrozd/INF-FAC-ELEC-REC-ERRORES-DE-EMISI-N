LPARAMETERS LPcTabla

LlPrimVez= .F.
LlSubTitulo= .F.
SELECT ReportXp
SET FILTER TO ALLTRIM(NomReporte)= ALLTRIM(UPPER(LPcTabla))
GO TOP
DO WHILE !EOF() AND !LlSubTitulo
   IF !EMPTY(SubTitulo)
      LlSubTitulo= .T.
   ENDIF
   IF !EOF()
      SKIP
   ENDIF
ENDDO
GO TOP
DO WHILE !EOF()
   IF LlSubTitulo AND !LlPrimVez
      .ROWS(ALLTRIM(STR(2))+":"+ALLTRIM(STR(2))).SELECT
      .SELECTION.INSERT
      LlPrimVez= .T.
   ENDIF
   IF LlSubTitulo
      LnOrden= Orden
      .cells(2,LnOrden).VALUE= ALLTRIM(SubTitulo)
      .cells(2,LnOrden).SELECT
      FPrLetra('Times New Roman',8,,)
      FBordes(1,-4138,-4105)
      .SELECTION.Interior.COLOR= RGB(226,239,218)
      .SELECTION.HorizontalAlignment= -4108
      .SELECTION.VerticalAlignment= -4160
      .SELECTION.WrapText= .T.
      .SELECTION.ORIENTATION= 0
      .SELECTION.AddIndent= .F.
      .SELECTION.IndentLevel= 0
      .SELECTION.ShrinkToFit= .F.
      .SELECTION.ReadingOrder= -5002
*!*	      .SELECTION.MergeCells= .T.      
   ENDIF
   IF !EOF()
      SKIP
   ENDIF
ENDDO
