LPARAMETERS LPcTabla,LPcNombreAr,LPcNombreFOX2X,LPcNombreXls,LPcNombreXlsPrin,LPcPassword, ;
	LPcNomHoja,LPcTitulos,LPcCampos,LPcCamSum,LPcCamDin,LPcForDin,LPcNomDin,LPcFAgreDin,LPlGrafica, ;
	LPnInserta,LPcEncabeza,LPlCreaObj,LPlCierraObj,LPlVisible,LPcMemos,LPcFAgreDat,LPnExpoCopy,LPlRejilla,LPcArchivoExistente

PUBLIC oObjeto
*WAIT WINDOW ' Generando Informe a Excel, Espere por favor.... ' NOWAIT


IF EMPTY(LPcNombreAr)
	LPcNombreAr= 'Hoja1'
ENDIF
IF EMPTY(GcDirectorio)
	GcDirectorio= 'c:\Informes Gerencial\'
ENDIF
IF EMPTY(LPnExpoCopy)
	LPnExpoCopy= 1
ENDIF
IF EMPTY(LPcNombreFOX2X)
	LPcNombreFOX2X= GcDirectorio+(LPcNombreAr)+'.dbf'

	IF !EMPTY(LPcTabla) AND EMPTY(LPcArchivoExistente)
		SELECT &LPcTabla
		IF !EMPTY(LPcCampos)
			SET COMPATIBLE OFF 	&& Estaba ON
			SET FIELDS TO &LPcCampos
		ENDIF
		AFIELDS(LmTipoCampo)
		SELECT &LPcTabla
		ERASE(LPcNombreFOX2X)

		DO CASE
			CASE LPnExpoCopy= 1
				IF !EMPTY(LPcCampos)
					COPY TO (LPcNombreFOX2X) TYPE FOX2X FIELDS &LPcCampos
				ELSE
					COPY TO (LPcNombreFOX2X) TYPE FOX2X
				ENDIF
			CASE LPnExpoCopy= 2
				IF !EMPTY(LPcCampos)
					EXPORT TO (LPcNombreFOX2X) TYPE XLS FIELDS &LPcCampos
				ELSE
					EXPORT TO (LPcNombreFOX2X) TYPE XLS
				ENDIF
			CASE LPnExpoCopy= 3
				IF !EMPTY(LPcCampos)
					COPY TO (LPcNombreFOX2X) TYPE XL5 FIELDS &LPcCampos
				ELSE
					COPY TO (LPcNombreFOX2X) TYPE XL5
				ENDIF
			OTHERWISE
				IF !EMPTY(LPcCampos)
					COPY TO (LPcNombreFOX2X) TYPE FOX2X FIELDS &LPcCampos
				ELSE
					COPY TO (LPcNombreFOX2X) TYPE FOX2X
				ENDIF
		ENDCASE

		SELECT &LPcTabla
		SET FIELDS TO ALL
		SET FIELDS OFF
		SET COMPATIBLE &GcGuardaSetCompatible
	ENDIF
ENDIF

IF EMPTY(LPcNomHoja)
	LPcNomHoja= LPcNombreAr
ENDIF

IF LPlCreaObj
	oObjeto= CREATE('Excel.Application')
	WITH oObjeto
		.Workbooks.OPEN(LPcNombreFOX2X)
		.Sheets.ADD
		.Sheets(LPcNombreAr).SELECT
		.Sheets(LPcNombreAr).NAME= LPcNomHoja
		.Sheets(LPcNomHoja).SELECT
		.Cells.SELECT
		*- Excel 97-2003=16384 - Excel 2003=65536 - Excel 2007=1048576 - Excel 2010=1048576

		*- Excel 97 = 8 
		*- Excel 2000 = 9
		*- Excel 2002 = 10
		*- Excel 2003 = 11
		*- Excel 2007 = 12
		*- Excel 2010 = 14
		*- Excel 2013 = 15

		FVersExc()
		IF GnVersExc< 12
			LnGuardaVer= -4143
			LPcNombreXls= GcDirectorio+(LPcNombreAr)+LcExtension
		ELSE
			LnGuardaVer= 51
			LPcNombreXls= GcDirectorio+(LPcNombreAr)+LcExtension
		ENDIF

		LlAbreArc= .F.
		DO WHILE !LlAbreArc
			IF FAbreArc(LPcNombreXls,)
				LlAbreArc= .T.
			ELSE
				MESSAGEBOX(' ¡Usted tiene ya el Archivo ['+LPcNombreXls+'] abierto!, ' +CHR(13)+ ;
					'por favor ciérrelo para poder continuar con este proceso. ' +CHR(13)+ ;
					'Presione [ENTER] cuando esté listo, ¡Gracias!',0+16+0,GcMensa)
			ENDIF
		ENDDO
		ERASE(LPcNombreXls)

		.APPLICATION.DisplayAlerts= .F.
		.APPLICATION.AlertBeforeOverwriting= .F.
		IF EMPTY(LPcPassword)
			.ActiveWorkbook.SAVEAS(LPcNombreXls,LnGuardaVer,,,,)
		ELSE
			.ActiveWorkbook.SAVEAS(LPcNombreXls,LnGuardaVer,LPcPassword,,,)
		ENDIF
		ERASE(LPcNombreFOX2X)
		FContenido(LPcTitulos,LPcCamSum,LPcCamDin,LPcForDin,LPcNomDin,LPcFAgreDin,LPlGrafica,LPnInserta,LPcEncabeza,LPcTabla,LPcMemos,LPcFAgreDat,LPcNomHoja,LPlRejilla,LPcArchivoExistente)
		.ActiveWorkbook.SAVE()
		.APPLICATION.DisplayAlerts= .T.
		.APPLICATION.AlertBeforeOverwriting= .T.
		.VISIBLE= LPlVisible
		IF LPlCierraObj
			FCierraObj(LPcNombreFOX2X,LPcNomHoja,LPcNombreXlsPrin,LPlVisible)
		ENDIF
	ENDWITH
ELSE
	WITH oObjeto
		.APPLICATION.DisplayAlerts= .F.
		.APPLICATION.AlertBeforeOverwriting= .F.
		IF !EMPTY(LPcTabla) AND EMPTY(LPcArchivoExistente)
			.ActiveWorkbook.SAVE()
			.Workbooks(LPcNombreXlsPrin).CLOSE
			.Workbooks.OPEN(LPcNombreFOX2X)
			.Cells.SELECT
			.SELECTION.COPY

			LlAbreArcX= .F.
			IF !EMPTY(LPcNombreXls)
				DO WHILE !LlAbreArcX
					IF FExis(LPcNombreXls,.F.)
						IF FAbreArc(LPcNombreXls,)
							FPausa(1000)
							LlAbreArcX= .T.
						ELSE
							FPausa(1000)
						ENDIF
					ELSE
						FPausa(1000)
					ENDIF
				ENDDO
			ENDIF
			.Workbooks.OPEN(LPcNombreXls)
		ENDIF
		IF	EMPTY(LPcArchivoExistente)
			.Workbooks(LPcNombreXlsPrin).ACTIVATE
			.Sheets.ADD
			.Sheets(.ActiveSheet.NAME).SELECT
			.Sheets(.ActiveSheet.NAME).NAME= LPcNomHoja
			.Sheets(LPcNomHoja).SELECT
		ELSE
			.Sheets(LPcArchivoExistente).SELECT
			.RANGE(LPcArchivoExistente).SELECT
		ENDIF

		IF !EMPTY(LPcTabla)
			IF	EMPTY(LPcArchivoExistente)
				.ActiveSheet.Paste
				.APPLICATION.CutCopyMode= .F.
				LcNombreFOX2X= LPcNombreFOX2X
				LPcNombreFOX2X= SUBSTR(LPcNombreFOX2X,AT_C('\',LPcNombreFOX2X,OCCURS('\',LPcNombreFOX2X))+1,LEN(LPcNombreFOX2X))
				.Workbooks(LPcNombreFOX2X).CLOSE
				ERASE(LcNombreFOX2X)
			ENDIF
			FContenido(LPcTitulos,LPcCamSum,LPcCamDin,LPcForDin,LPcNomDin,LPcFAgreDin,LPlGrafica,LPnInserta,LPcEncabeza,LPcTabla,LPcMemos,LPcFAgreDat,LPcNomHoja,LPlRejilla,LPcArchivoExistente)
		ELSE
			IF	!EMPTY(LPcArchivoExistente)
				FContenido(LPcTitulos,LPcCamSum,LPcCamDin,LPcForDin,LPcNomDin,LPcFAgreDin,LPlGrafica,LPnInserta,LPcEncabeza,LPcTabla,LPcMemos,LPcFAgreDat,LPcNomHoja,LPlRejilla,LPcArchivoExistente)
			ENDIF
			IF !EMPTY(LPcFAgreDat) AND EMPTY(LPcArchivoExistente)
				&LPcFAgreDat(LPcNombreXlsPrin,LPcNomHoja,LPnInserta,LPcEncabeza)
			ENDIF
		ENDIF
		.Workbooks(LPcNombreXlsPrin).ACTIVATE
		.ActiveWorkbook.SAVE()
		.APPLICATION.DisplayAlerts= .T.
		.APPLICATION.AlertBeforeOverwriting= .T.
		.VISIBLE= LPlVisible
		IF LPlCierraObj
			FCierraObj(LPcNombreFOX2X,LPcNomHoja,LPcNombreXlsPrin,LPlVisible)
		ENDIF
	ENDWITH
ENDIF


FUNCTION FContenido(LPcTitulos,LPcCamSum,LPcCamDin,LPcForDin,LPcNomDin,LPcFAgreDin,LPlGrafica,LPnInserta,LPcEncabeza,LPcTabla,LPcMemos,LPcFAgreDat,LPcNomHoja,LPlRejilla,LPcArchivoExistente)
	IF	EMPTY(LPcArchivoExistente)
		LnCombina2= 0
		LnTotalCol= ALEN(LmTipoCampo,1)
		.Cells.SELECT
		.SELECTION.COLUMNS.AUTOFIT
		.SELECTION.COLUMNS.AutoFilter
		.APPLICATION.CutCopyMode= .F.
		IF LPlRejilla
			LcHoja= STRTRAN(STRTRAN(LPcNomHoja, ' ', '_'),'-','_')
			LnFil= .ActiveSheet.UsedRange.ROWS.COUNT
			.RANGE(.RANGE(.Cells(1,1),.Cells(LnFil, LnTotalCol)).Address()).SELECT
			.ActiveSheet.ListObjects.ADD.NAME= LcHoja
			.ActiveSheet.ListObjects(LcHoja).TableStyle= "TableStyleMedium21"
			.ActiveSheet.ListObjects(LcHoja).ShowTotals= .T.
		ENDIF
	ENDIF
	IF !EMPTY(LPcCamDin)
		FDinamic(LPcCamDin,LPcForDin,LPcNomDin,LPcFAgreDin,LPlGrafica,LPnInserta,LPcArchivoExistente)
	ENDIF
	IF	EMPTY(LPcArchivoExistente)
		.Sheets(LPcNomHoja).SELECT
		LnFila= 1
		.Cells.SELECT
		FPrLetra('Times New Roman',8,,)
		IF EMPTY(LPcTitulos)
			IF  USED('ReportXp')
				SELECT ReportXp
				COUNT FOR ALLTRIM(NomReporte)= ALLTRIM(UPPER(LPcTabla)) TO LnCont
				IF LnCont>0
					COUNT FOR ALLTRIM(NomReporte)= ALLTRIM(UPPER(LPcTabla)) AND !EMPTY(Combina2) TO LnCombina2
					IF !LPlRejilla
						FSumasTa(LPcTabla)
					ELSE
						LcHoja= STRTRAN(STRTRAN(LPcNomHoja, ' ', '_'),'-','_')
						FSumasTG(LcHoja,LPcTabla)
					ENDIF
					FTitulTa(LPcTabla)
					FSubTitu(LPcTabla)
					FExForTa(LnTotalCol,LPcTabla)
					FComenta(LPcTabla)
					LnOcurre= FCombina(LPcTabla,'Combina1')
					IF LnOcurre=1
						.RANGE(.Cells(1,1), .Cells(1, LnTotalCol)).SELECT
						FPrLetra('Bookman Old Style',10,.T.,)
						IF LnCombina2>0
							FOR LnCont=1 TO 3
								.SELECTION.BORDERS(LnCont).LineStyle= 1
								.SELECTION.BORDERS(LnCont).Weight= -4138
								.SELECTION.BORDERS(LnCont).ColorIndex= -4105
							ENDFOR
						ELSE
							FBordes(1,-4138,-4105)
							FRelleno(8,-4105)
						ENDIF
						LnFila= LnFila+1
						LPnInserta= LPnInserta+1
					ENDIF
					LnOcurre= FCombina(LPcTabla,'Combina2')
					IF LnOcurre=1
						.RANGE(.Cells(1,1), .Cells(1, LnTotalCol)).SELECT
						FPrLetra('Arial Black',12,.T.,)
						FBordes(1,-4138,-4105)
						.SELECTION.Interior.COLOR= RGB(169,208,142)
						LnFila= LnFila+1
						LPnInserta= LPnInserta+1
					ENDIF
				ENDIF
			ENDIF
		ELSE
			LnTotalCol= FExTitul(LPcTitulos)
		ENDIF

		LnTotalCol= ALEN(LmTipoCampo,1)
		.RANGE(.Cells(LnFila,1), .Cells(LnFila, LnTotalCol)).SELECT
		FExForCo(LnTotalCol)
		FPrLetra('Tahoma',10,.T.,2)
		.Cells.SELECT
		.SELECTION.COLUMNS.AUTOFIT
		IF !EMPTY(LPcMemos)
			FMemos(LnTotalCol,LPcTabla,LPcMemos,LPnInserta,LnFila)
			.Cells.SELECT
			.SELECTION.ROWS.AUTOFIT
		ENDIF
		.RANGE(.Cells(LnFila,1), .Cells(LnFila, LnTotalCol)).SELECT
		IF USED('ReportXp')
			SELECT ReportXp
			COUNT FOR ALLTRIM(NomReporte)= ALLTRIM(UPPER(LPcTabla)) TO LnCont
			IF LnCont>0
				FExForTa(LnTotalCol,LPcTabla)
			ELSE
				FExForCo(LnTotalCol)
			ENDIF
		ELSE
			FExForCo(LnTotalCol)
		ENDIF
		FPrLetra('Times New Roman',8,.T.,2)
		.ROWS("1:"+ALLTRIM(STR(LPnInserta))).SELECT
		.SELECTION.INSERT
		.RANGE(.Cells(LPnInserta+LnFila,1), .Cells(LPnInserta+LnFila, LnTotalCol)).SELECT
		IF LnCombina2> 0
			FOR LnCont= 1 TO 4
				IF LnCont<>3
					.SELECTION.BORDERS(LnCont).LineStyle= 1
					.SELECTION.BORDERS(LnCont).Weight= -4138
					.SELECTION.BORDERS(LnCont).ColorIndex= -4105
				ENDIF
			ENDFOR
		ELSE
			FBordes(1,-4138,-4105)
		ENDIF
		.SELECTION.Interior.COLOR= RGB(0,125,0)
		IF !EMPTY(LPcCamSum)
			FSumaCol(LnTotalCol,LPcCamSum,LPnInserta)
		ENDIF
	ENDIF
	IF !EMPTY(LPcFAgreDat)
		&LPcFAgreDat(LPcNomHoja,LPnInserta,LPcEncabeza)
	ENDIF
	IF	EMPTY(LPcArchivoExistente)
		.Sheets(LPcNomHoja).SELECT
		.RANGE("A1").SELECT
		IF !EMPTY(LPcEncabeza)
			FEncabez(LPcEncabeza)
		ENDIF
	ENDIF
	.ActiveWindow.ScrollWorkbookTabs(0)
	.Sheets(1).ACTIVATE
	.RANGE("A1").SELECT
ENDFUNC

FUNCTION FCierraObj(LPcNombreFOX2X,LPcNomHoja,LPcNombreXlsPrin,LPlVisible)
	IF !EMPTY(LPcNombreXlsPrin)
		.Workbooks(LPcNombreXlsPrin).ACTIVATE
	ENDIF
	.APPLICATION.DisplayAlerts= .F.
	.APPLICATION.AlertBeforeOverwriting= .F.
	.ActiveWorkbook.SAVE()
	.Sheets(LPcNomHoja).SELECT
	.Sheets(LPcNomHoja).ACTIVATE
	.RANGE("A1").SELECT
	ERASE(LPcNombreFOX2X)
	IF !LPlVisible
		.ActiveWindow.CLOSE
		.APPLICATION.DisplayAlerts= .T.
		.APPLICATION.AlertBeforeOverwriting= .T.
		.APPLICATION.QUIT
	ENDIF
ENDFUNC
