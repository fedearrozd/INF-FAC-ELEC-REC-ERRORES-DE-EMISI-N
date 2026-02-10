# INF-FAC-ELEC-REC-ERRORES-DE-EMISI-N
INFORME DE FACTURACIÓN ELECTRÓNICA – RECHAZOS – ERRORES DE EMISIÓN

Se desarrolló informe consolidado de los documentos electrónicos que viajaron a la DIAN desde el ingreso
de vigencia de la resolución 165 del 2023 y que por algún momento fueron rechazados. Este informe se
envía a todos los directores y seccionales de forma automática, cada hora, y no envía nada si no hay
rechazos. Envía además a algunos usuarios de oficina principal el consolidado de todas las seccionales
que hayan presentado rechazos, indicando en la última columna el motivo del rechazo.

|Áreas Beneficiadas| Sistemas.|
| :--- | :--- |
|Ubicación de Fuentes |\\194.168.0.65\bkdesarrollo\FEDEARROZ\Facturación|
|Ejecutable |Descripción|
|FFacErrEmiMin.exe |Genera Informe de facturación electrónica errores de emisión de los documentos electrónicos como facturas, notas y otros que por algún motivo fueron rechazados por la DIAN. En el consolidado indica el motivo del rechazo en las últimas columnas. Diseñado para programación de horas y minutos.|
|FFacErrEmiDia.exe |Genera Informe de facturación electrónica errores de emisión para programación diaria.|
|FFacErrEmiSem.exe |Genera Informe de facturación electrónica errores de emisión para programación semanal.|

Ante perdida o reconfiguracion de la base de datos configurar el archivo fconfsql.dbf con las conexiones y credenciales.
