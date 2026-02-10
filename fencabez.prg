LPARAMETERS LPcEncabeza

LnTama= 14
LnIni= 1
LnOcurre= OCCURS(',',LPcEncabeza)
IF LnOcurre>=0
   IF LnOcurre>0
      LnFin= ATC(',',LPcEncabeza,1)-LnIni
   ENDIF
   LnOcurre=LnOcurre+1
   FOR LnCnt=1 TO LnOcurre
      IF LnCnt=LnOcurre
         LnFin= LEN(LPcEncabeza)-(LnIni-1)
      ENDIF
      LcValor1= ALLTRIM(SUBSTR(LPcEncabeza,LnIni,LnFin))
      .ROWS(ALLTRIM(STR(LnCnt))+":"+ALLTRIM(STR(LnCnt))).SELECT
      .SELECTION.INSERT
      .cells(LnCnt,1).VALUE=LcValor1
      LnTama= LnTama-1
      FPrLetra('Times New Roman',LnTama,.T.,10)
      LnIni= ATC(',',LPcEncabeza,LnCnt)+1
      LnFin= ATC(',',LPcEncabeza,LnCnt+1)-LnIni
   ENDFOR
ENDIF

RETURN LnOcurre
