LPARAMETERS LPcPrograma, LPcParametro, LPcMensaErr, LPcMensaEr2, LPcCodigo, LPcNumLinea, LPcAlias, LPcTabla

GuardarAlias= ALIAS()

IF TYPE('LPcMensaEr2')= 'N'
   LPcMensaEr2= ALLTRIM(STR(LPcMensaEr2,20,2))
ENDIF
IF TYPE('LPcNumLinea')= 'N'
   LPcNumLinea= ALLTRIM(STR(LPcNumLinea,20,2))
ENDIF
IF TYPE('LPcCodigo')= 'N'
   LPcCodigo= ALLTRIM(STR(LPcCodigo,20,2))
ENDIF
IF EMPTY(LPcTabla)
   LPcTabla= ''
ENDIF

*!*	LnMensaje= MESSAGEBOX( ;
*!*	   [Program : ]+ LPcPrograma +CHR(13)+ ;
*!*	   [Parámetr: ]+ LPcParametro +CHR(13)+ ;
*!*	   [Error   : ]+ LPcMensaErr +CHR(13)+ ;
*!*	   [Mensaje : ]+ LPcMensaEr2 +CHR(13)+ ;
*!*	   [Codigo  : ]+ LPcCodigo +CHR(13)+ ;
*!*	   [NumLinea: ]+ LPcNumLinea +CHR(13)+ ;
*!*	   [Alias   : ]+ LPcAlias +CHR(13)+ ;
*!*	   [Tabla   : ]+ LPcTabla +CHR(13) ,1+16+0,'Mensaje de Error...')

IF !FILE('FeErrSis.dbf')
   CREATE TABLE FeErrSis(Fecha D(8), Hora c(5), Programa c(100), Parametro c(50), Camino c(100), MensaErr c(100), MensaEr2 c(100), Codigo c(10), NumLinea c(5), LAlias c(20), Tabla c(200), Maquina c(50), UsuarioRed c(50))
ELSE
   USE FeErrSis.DBF IN 0 SHARED
ENDIF

LcMaquina= UPPER(SUBSTR(SYS(0),1,AT('#',SYS(0))-1))
LcUsuarioRed= UPPER(SUBSTR(SYS(0),AT('#',SYS(0))+2))
INSERT INTO FeErrSis(Fecha,Hora,Programa,Parametro,Camino,MensaErr,MensaEr2,Codigo,NumLinea,LAlias,Tabla,Maquina,UsuarioRed) VALUES(DATE(),TIME(),LPcPrograma,LPcParametro,GCamino,LPcMensaErr,LPcMensaEr2,LPcCodigo,LPcNumLinea,LPcAlias,LPcTabla,LcMaquina,LcUsuarioRed)
USE IN SELECT('FeErrSis')

IF !EMPTY(GuardarAlias)
   SELECT (GuardarAlias)
ENDIF

RETURN .F.