Sub DemandaRepetidaFrete()
	Dim oDoc As Object
	Dim oSheet As Object
	
	Dim Cell6 As String
	Dim Cell21 As String
	
	'Obter a planilha ativa
	oDoc = ThisComponent
	oSheet = oDoc.Sheets.getByName("Planilha1")
	
	Cell6 = oSheet.getCellRangeByName("B6").getString
	Cell21 = oSheet.getCellRangeByName("B21").getString

	MsgBox "Esse erro ocorreu nessas células:" & Chr(10) & Cell6 & Chr(10) & Cell21
		
End Sub

