LPARAMETERS LPnTotalCol, LPcTabla, LPcMemos, LPnInserta, LPnFila

LnCol= .ActiveSheet.UsedRange.COLUMNS.COUNT
LnFil= .ActiveSheet.UsedRange.ROWS.COUNT
LnIni= 1
LnOcurre= OCCURS(',',LPcMemos)
IF LnOcurre>=0
	IF LnOcurre>0
		LnFin= ATC(',',LPcMemos,1)-LnIni
	ENDIF
	LnOcurre=LnOcurre+1
	FOR LnCnt=1 TO LnOcurre
		IF LnCnt=LnOcurre
			LnFin= LEN(LPcMemos)-(LnIni-1)
		ENDIF
		LcCampo= ALLTRIM(SUBSTR(LPcMemos,LnIni,LnFin))
		LnColumna= FMatrizFor(LPnTotalCol,LcCampo)
		IF LnColumna> 26
			LcColSuma= SUBSTR(STRTRAN(.Cells(LPnInserta+1,LnColumna).Address(),'$',''),1,2)
		ELSE
			LcColSuma= SUBSTR(STRTRAN(.Cells(LPnInserta+1,LnColumna).Address(),'$',''),1,1)
		ENDIF
		LnColMemo= LcColSuma+':'+LcColSuma
		LcCamMemo= LPcTabla+'.'+LcCampo
		IF !EMPTY(LPnFila)
			LcRango= STRTRAN(.Cells(LPnInserta,LnColumna).Address(),'$','')+':'+FDevCamp(.Cells(1,LnColumna).Address(),2,'$')+ALLTRIM(STR(LnFil))
		ELSE
			LcRango= ''
		ENDIF
		FAjuText(LnColMemo,1,-4107,.T.,0,.F.,0,.F.,-5002,.F.,LcRango)
		FAjuACol(LnColMemo,100)
		SELECT (LPcTabla)
		GO TOP
		FOR LnLin=1 TO LnFil
			LnFilMemo= LcColSuma+ALLTRIM(STR(LnLin+LPnFila))
			.RANGE(LnFilMemo).SELECT
			.RANGE(LnFilMemo).VALUE= (&LcCamMemo)
			SELECT (LPcTabla)
			IF !EOF()
				SKIP
			ENDIF
		ENDFOR
		LnIni= ATC(',',LPcMemos,LnCnt)+1
		LnFin= ATC(',',LPcMemos,LnCnt+1)-LnIni
	ENDFOR
ENDIF
