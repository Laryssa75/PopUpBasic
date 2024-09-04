Sub DemandaRepetidaComplemento()
	Dim oDoc As Object
	Dim oSheet As Object
	
	Dim Cell15 As String
	Dim Cell24 As String
	
	'Obter a planilha ativa
	oDoc = ThisComponent
	oSheet = oDoc.Sheets.getByName("Planilha1")
	
	Cell15 = oSheet.getCellRangeByName("B15").getString
	Cell24 = oSheet.getCellRangeByName("B24").getString

	MsgBox "Esse erro ocorreu nessas células" & Chr(10) & Cell15 & Chr(10) & Cell24
		
End Sub
