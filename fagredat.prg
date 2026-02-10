LPARAMETERS LPnTotalCol,LPnTotalFil,LPnInserta,LPcCampoCol,LPnFilas,LPcDetalle,LPlLetra,LPlBordes,LPlRelleno

.Sheets("Error_Emision").SELECT
.Sheets('Error_Emision').ACTIVATE
.ActiveWorkbook.Sheets("Error_Emision").TAB.COLOR= 5296274
.Sheets("Error_Emision").MOVE(.Sheets(1))
.ActiveWindow.DisplayGridlines= .F.
.Sheets("Hoja1").SELECT
.Sheets('Hoja1').ACTIVATE
.APPLICATION.DisplayAlerts= .F.
.APPLICATION.AlertBeforeOverwriting= .F.
.ActiveWindow.SelectedSheets.DELETE
.RANGE("B:Q").SELECT
.COLUMNS("B:Q").EntireColumn.AutoFit
.RANGE("Error_Emision").SELECT
.ActiveSheet.ListObjects("Error_Emision").NAME = "Errores_Emision"
.SELECTION.FormatConditions.ADD(2, 6, "=$E4< HOY()")
.SELECTION.FormatConditions(.SELECTION.FormatConditions.COUNT).SetFirstPriority
.SELECTION.FormatConditions(1).FONT.ThemeColor = 1
.SELECTION.FormatConditions(1).FONT.TintAndShade = 0
.SELECTION.FormatConditions(1).Interior.PatternColorIndex = 0
.SELECTION.FormatConditions(1).Interior.COLOR = 192
.SELECTION.FormatConditions(1).Interior.TintAndShade = 0
.SELECTION.FormatConditions(1).Interior.PatternTintAndShade = 0
.SELECTION.FormatConditions.ADD(2, 6, "=$E4>= HOY()")
.SELECTION.FormatConditions(.SELECTION.FormatConditions.COUNT).SetFirstPriority
.SELECTION.FormatConditions(1).FONT.COLOR = -16777024
.SELECTION.FormatConditions(1).FONT.TintAndShade = 0
.SELECTION.FormatConditions(1).Interior.PatternColorIndex = 0
.SELECTION.FormatConditions(1).Interior.ThemeColor = 3
.SELECTION.FormatConditions(1).Interior.TintAndShade = -0.249946592608417
.SELECTION.FormatConditions(1).Interior.PatternTintAndShade = 0
.SELECTION.FormatConditions(1).StopIfTrue = .F.
.ActiveSheet.ListObjects("Errores_Emision").ListColumns("EMAIL").TotalsCalculation= 3
.RANGE("A1").SELECT
