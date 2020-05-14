Sub Main
	Call IndexedExtraction()	'Ejemplo-Detalle de ventas.IMD
End Sub


' Datos: Extracción indexada
Function IndexedExtraction
	Const WI_IE_NUMFLD  = 1
	Const WI_IE_CHARFLD = 2
	Const WI_IE_TIMEFLD = 3
	
	Set db = Client.OpenDatabase("Ejemplo-Detalle de ventas.IMD")
	Set task = db.IndexedExtraction
	task.IncludeAllFields
	task.FieldToUse =  "COD_PROD"
	task.FieldValueIs2 WI_IE_GTEQUAL, "04", WI_IE_CHARFLD
	task.CreateVirtualDatabase = False
	dbName = "Extraccion_Index_01.IMD"
	task.OutputFilename =  dbName
	task.PerformTask
	Set task = Nothing
	Set db = Nothing
	Client.OpenDatabase (dbName)
End Function