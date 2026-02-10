LPARAMETERS LPcTabla,LPcPerfil,LPcPassword,LPcAsunto,LPcRutaAdjuntos,LPnPuerto

PUBLIC oMAIL, oConf, GmAError
DIMENSION GmAError[7,2]
LnGuardaSetReprocess= SET('REPROCESS')
SET REPROCESS TO -1
SET MEMOWIDTH TO 86

LcPerfil= LPcPerfil
LcPassword= LPcPassword
LcAsunto= LPcAsunto
LnImportancia= 1		&& 0, 1, 2
LnPrioridad= 1			&& -1, 0, 1
LlExisArch= .T.
LlExitoso= .F.
LcCFG = "http://schemas.microsoft.com/cdo/configuration/"
oMAIL = CREATEOBJECT("CDO.Message")
oConf = CREATEOBJECT("CDO.Configuration")

SELECT &LPcTabla
WITH oConf
	.FIELDS.ITEM(LcCFG+"sendusing")= 2
	.FIELDS.ITEM(LcCFG+"smtpserver")= "smtp.office365.com"
	.FIELDS.ITEM(LcCFG+"smtpserverport")= LPnPuerto
	.FIELDS.ITEM(LcCFG+"smtpauthenticate")= 1
	.FIELDS.ITEM(LcCFG+"smtpconnectiontimeout")= 60
	.FIELDS.ITEM(LcCFG+"sendusername")= LcPerfil
	.FIELDS.ITEM(LcCFG+"sendpassword")= LcPassword
	.FIELDS.ITEM(LcCFG+"smtpusessl")= .T.
	.FIELDS.UPDATE()
ENDWITH

WITH oMAIL
	.Configuration = oConf
	.FROM = LcPerfil
	.TO = (Destinatar)
	.Cc= (CopiaA)
	.BCC= (CopiaOcult)
	.Subject= LcAsunto
	*.Importance= LnImportancia
	*.priority= LnPrioridad
	.TextBody = Mensaje

	LnNumLin= 1
	LnNumLinMemo= MEMLINES(Adjuntos)
	IF (LnNumLinMemo= 0)
		LnNumLinMemo= 1
	ENDIF

	DO WHILE (LnNumLin <= LnNumLinMemo)
		LmAdjuntos= MLINE(UPPER(Adjuntos), LnNumLin)
		LnNumLin= LnNumLin + 1
		LcArchAdjunto= GcDirectorio+(LmAdjuntos)

		IF FILE(LcArchAdjunto)
			.AddAttachment(LcArchAdjunto)
			FPausa(1000)
			LlExisArch= .T. AND LlExisArch
		ELSE
			WAIT WINDOWS ('¡El archivo : '+CHR(13)+'['+GcDirectorio+(LmAdjuntos)+']'+CHR(13)+CHR(13)+ ;
				'!NO EXISTE!.  ¡NO SERÁ ENVIADO!.'+CHR(13)+CHR(13)+ 'Deberá crearlo nuevamente para enviarlo.') NOWAIT
			LlExisArch= .F. AND LlExisArch
		ENDIF
	ENDDO

	IF LlExisArch
		.SEND
	ENDIF

	=AERROR(GmAError)

	IF MESSAGE(1)= '.SEND'
		LlExitoso= .F.
	ELSE
		LlExitoso= .T.
	ENDIF
ENDWITH

RELEASE oMAIL,oConf

SET REPROCESS TO LnGuardaSetReprocess

RETURN LlExitoso
