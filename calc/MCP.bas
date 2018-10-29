REM  *****  BASIC  *****

Sub Main

	Dim oDoc as Object
	Dim oSheet as Object
	Dim shName as String
	Dim oCell as Object
	Dim oRange as Object
	Dim rowNb as Integer
	Dim colNb as Integer
	Dim oSearch As Object
	Dim oReplace As Object
	Dim oResult As Object

	oDoc = ThisComponent
	shName="Arkusz1"
	oSheet = oDoc.sheets.getByName(shName)
	ThisComponent.getCurrentController.select((oRange)
	
End Sub


' Jakieś dziwne kodowanie polskich liter przez "MAT"
' zamienia na poprawne polskie litery.
Sub MatCode()

	Dim oSheet as Object
	Dim shName as String
	
	oDoc = ThisComponent
	shName="Arkusz1"
	oSheet = oDoc.sheets.getByName(shName)
		
	before = array("NQ", "NCe", "NCE", "N3", "N#", "NDn", "NDN", "NBo", "NBO", "Nz", ")?", "N?", "*?")
	after =  array( "ą",  "ę",    "Ę",  "ł",  "Ł",  "ń",   "Ń",   "ó",   "Ó",   "ś",  "ź",  "ż",  "Ż") 
	ReplaceAll (before, after, oSheet)

End Sub

' Na arkuszu oSheet zamienia Stringi podane w tablicy before na Stringi z tablicy after
Sub ReplaceAll (before, after, oSheet)

	Dim oReplace As Object
	Dim n As Long
	
	oReplace = oSheet.createReplaceDescriptor
	For n = lbound(before()) To ubound(before())
		oReplace.SearchString = before(n)
		oReplace.ReplaceString = after(n)
		oSheet.replaceAll(oReplace)
	Next n
End Sub

' W bieżącym skoroszycie amienia wszystkie wystąpienia "oldTxt" na newTxt
' ( na wszystkich arkuszach skoroszytu)
Sub ReplaceText(oldTxt, newTxt )
	Dim oDoc As Object
	Dim oSheet As Object
	Dim oReplaceDescriptor As Object
	Dim i As Integer
	
	oDoc = ThisComponent
	oSheet = oDoc.Sheets(0)
	
	ReplaceDescriptor = oSheet.createReplaceDescriptor()
	ReplaceDescriptor.SearchString = oldTxt
	ReplaceDescriptor.ReplaceString = newTxt
	ReplaceDescriptor.SearchCaseSensitive = True
	For i = 0 to oDoc.Sheets.Count - 1
	   oSheet = oDoc.Sheets(i)
	   oSheet.ReplaceAll(ReplaceDescriptor)
	Next i

End Sub

' Zaznacza wszystkie komórki zawierające szukany tekst.
' i wypisuje ich adresy
Sub ShowFound()
	Dim oDoc as Object
	Dim oSheet as Object
	Dim shName as String
	Dim oRange as Object
	Dim oSearch As Object, oResult As Object

	oDoc = ThisComponent
	shName="Arkusz1"
	oSheet = oDoc.sheets.getByName(shName)
	
	oSearch = oSheet.createSearchDescriptor
	oSearch.SearchString = "^Kac[a-z]"
	oSearch.SearchRegularExpression = TRUE
	oSearch.SearchCaseSensitive = TRUE
	oResult =  oSheet.findAll(oSearch)
	
	'cellsAddress = split(oResult.AbsoluteName, ";")
	For n =  0  To  oResult.count(elementNames) -1
		oRange = oResult(n)
		ThisComponent.getCurrentController.select((oRange)
		print oRange.AbsoluteName
	next

End Sub

' Przypisuje pracownikowi komórkę organizacyjną
Sub WhatUnit()

	Dim oDoc as Object
	Dim oSheet as Object, dataSheet as Object
	Dim shName as String
	Dim oCell as Object
	Dim rowNb as Integer
	Dim colNb as Integer
	Dim name as String

	oDoc = ThisComponent
	shName="NotNominalTime"
	oSheet = oDoc.sheets.getByName(shName)
	
	shName="Workers"
	dataSheet = oDoc.sheets.getByName(shName)
		
	for rowNb= 1 to 510
		searchName = oSheet.getCellByPosition(1,rowNb).getString
		searchName = Format(searchName, "&gt;")
		
		' Szukam po numerze ewidencyjnym
		searchNb = oSheet.getCellByPosition(0,rowNb).getString
		
		'szukam numeru ewidencyjnego na arkuszu "Pracownicy"		
		for rowNb1= 2 to 510
			firstName = dataSheet.getCellByPosition(3,rowNb1).getString
			surName =   dataSheet.getCellByPosition(2,rowNb1).getString
			found = surName & " " &  firstName
			foundName = Format(found, "&gt;")
			foundNb = dataSheet.getCellByPosition(0,rowNb1).getString
			
			if searchNb = foundNb then 
				unit = dataSheet.getCellByPosition(8,rowNb1).getString
				oSheet.getCellByPosition(3,rowNb).String = unit
			endif
		next rowNb1
	next rowNb


End Sub

