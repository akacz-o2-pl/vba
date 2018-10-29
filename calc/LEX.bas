REM  *****  BASIC  *****

Sub Main

End Sub

' Zwraca informację o komórkach z czerwonym tekstem
Sub findRed
	Dim oDoc as Object
	Dim oSheet as Object
	Dim shName as String
	Dim oCell as Object	
	Dim rowNb as Integer
	Dim colNb as Integer
	
	dim oSelect as Object
	dim oCurs
	dim lastRow as Integer
	dim lastCol as Integer

	oDoc = ThisComponent  
	oSheet= oDoc.getCurrentController.activeSheet
	
	' Wybiera obszar wokół aktywnej komórki:
	oCell = oSheet.getCellByPosition(1,1)
	oDoc.getCurrentController.select( oCell)
	
	oSelect = oDoc.CurrentController.Selection   
  	oCurs =oSheet.createCursorByRange(oSelect)
    oCurs.collapseToCurrentRegion
    'oDoc.CurrentController.Select(oCurs) 
    
     oRange=oCurs
     oDoc.CurrentController.Select(oRange) 
    
	lastRow=oCurs.getRows().Count
	lastCol=oCurs.getColumns().Count
			
	for colNb = 0 to lastCol-1
		For rowNB = 1 to lastRow-1
		
			oCell = oSheet.getCellByPosition(colNb, rowNb)


			if oCell.CharColor = 16711680 AND oCell.getString <> "" then  
				oRow=oRange.Rows(rowNb)
				oDoc.getCurrentController.select(oCell)	
				print "wiersz nr: " & rowNb & "kolumna " & colNb
			endif
		Next rowNB
	next colNb
      
End Sub




' Generuje na arkuszu "insert" w wierszu o podanym numerze  zapytanie sql wstawiające dane z podanego wiersza akrusza "pk"
Sub addPK

	Dim oDoc as Object
	Dim oSheet as Object
	Dim shName as String
	Dim oCell as Object	
	Dim rowNb as Integer
	Dim colNb as Integer
	Dim q as String
	Dim v as String
	
	' Workbook:pk_wykaz.xls
	oDoc = ThisComponent               
	shName="pk"
	oSheet = oDoc.sheets.getByName(shName)
	
	'	Wybór wiersza
	rowNb = InputBox("Podaj nr wiersza")
	rowNb = rowNb -1
	q = "INSERT INTO `polecenia_komendanta` ( `id` , `nr` , `kom_org`,  `rok` , `z_dnia` , `w_sprawie` , `dokument` , `active`, `zal_name` , `zal` , `zmienione_przez` , `uchylone_przez` ) VALUES (NULL, "
	                   
	for colNb= 1 to 11
		oCell=oSheet.getCellByPosition(colNb,rowNb)   
		v = oCell.getString
		
		if colNb=9 then
			v=oSheet.getCellByPosition(colNb-1,rowNb).getString
		endif
		
		    
		
		if v = "" then 
			q = q & " " & "NULL" & ", "
		else 
				q = q & " '" & v & "', "
		endif
	next colNb
	q = left(q, Len(q)-2) & ");"
	
	' Zapisanie kwerendy na arkuszu insert
	shName="insert"
	oSheet = oDoc.sheets.getByName(shName)
	oCell=oSheet.getCellByPosition(0,rowNb)   
	oCell.setString(q)
	ThisComponent.getCurrentController.select((oCell)
	
End Sub


' Generuje na arkuszu "insert" w wierszu o podanym numerze  zapytanie sql wstawiające dane z podanego wiersza akrusza "zk"
Sub addZK

	Dim oDoc as Object
	Dim oSheet as Object
	Dim shName as String
	Dim oCell as Object	
	Dim rowNb as Integer
	Dim colNb as Integer
	Dim q as String
	Dim v as String
	
	' Workbook: zk_wykaz.xls
	oDoc = ThisComponent               
	shName="zk"
	oSheet = oDoc.sheets.getByName(shName)
	
	' Wybór wiersza
	rowNb = InputBox("Podaj nr wiersza")
	rowNb = rowNb -1
              
	q = "INSERT INTO `zarzadzenia_komendanta` ( `id` , `nr` , `kom_org`,  `rok` , `z_dnia` , `w_sprawie` , `dokument` , `active`, `zal` , " 
	q = q & "`zmienione_przez` , `uchylone_przez`, `wykonano` ) VALUES (NULL, "
               
	for colNb= 1 to 11
		oCell=oSheet.getCellByPosition(colNb,rowNb)       
		v = oCell.getString
		if v = "" then 
			q = q & " " & "NULL" & ", "
		else 
				q = q & " '" & v & "', "
		endif
	next colNb
	q = left(q, Len(q)-2) & ");"
	
	' Zapisanie kwerendy na arkuszu insert
	shName="insert"
	oSheet = oDoc.sheets.getByName(shName)
	oCell=oSheet.getCellByPosition(0,rowNb)   
	oCell.setString(q)
	ThisComponent.getCurrentController.select((oCell)
	
End Sub


REM Tworzy sql do aktualizacji aktu prawnego
Sub Prawo_Update

	Dim oDoc as Object
	Dim oSheet as Object
	Dim shName as String
	Dim oCell as Object	
	Dim rowNb as Integer
	Dim colNb as Integer
	Dim q as String
	Dim v as String

	' Workbook: prawo_wykaz.xls
	oDoc = ThisComponent               
	shName="Arkusz1"
	oSheet = oDoc.sheets.getByName(shName)
	
	rowNb = InputBox("Podaj nr wiersza")
	rowNb = rowNb -1
	search = oSheet.getCellByPosition(6,rowNb).getString
	newDoc = oSheet.getCellByPosition(16,rowNb).getString
	newPub  = oSheet.getCellByPosition(17,rowNb).getString

sql = "UPDATE `akty_prawne` SET `dokument` = """& _
      newDoc & """, `publikator` = """& newPub &  """ WHERE `dokument` = """& search & """"


print sql

End Sub


' Zaznacza wiersze dla których tekst w kolumnie wykonane ma czaerwony kolor
Sub Prawo_Wykonane
	
	Dim i As Integer
	Dim oDoc as Object
	Dim oSheets
	Dim oSheet
	Dim oCell
	Dim s As String, result as String
	
	' Workbook: prawo_wykaz.xls, zk_wykaz.xls, pk_wykaz.xls
	oDoc = ThisComponent               
	shName="zk"
	oSheet = oDoc.sheets.getByName(shName)

	For i = 0 to 1566
		oCell = oSheet.getCellByPosition(11,i) ' GetCell L, i
		s = oCell.getString
		if s  = "1" AND   oCell.CharColor = 16711680 then  
			oRow=oSheet.Rows(i)
			ThisComponent.getCurrentController.select( oRow)	
			result =  oSheet.getCellByPosition(6,i).String
			print result
		endif
	Next

End Sub

' Skoroszyt wykaz aktów prawnych.ods
' Przygotowuje arkusz1 do legalis
Sub Lex2legalis
	Dim i As Integer
	Dim oDoc as Object
	Dim oSheets
	Dim oSheet
	Dim oRange
	Dim oCell
	Dim lastRow as Integer
	Dim s As String, result as String
	
	oDoc = ThisComponent    
	if oDoc.Title <> "wykaz aktów prawnych.ods" then exit sub
	
	           
	shName="Arkusz1"
	oSheet = oDoc.sheets.getByName(shName)
	lastRowNb =204
	oRange= oSheet.getCellRangeByName("A2:J204")
	

	result=""
	for rowNb = 0 to lastRowNb-2
		title = oRange.getCellByPosition(6,rowNb).String
		typ =  oRange.getCellByPosition(5,rowNb).String

		if instr(typ, "ustawa")  = 1 then
			result = "Ustawa " & title
		endif
		
		if instr(typ, "rozp.") = 1 then
			result = "Rozporządzenie" & title
		endif
		
		ThisComponent.getCurrentController.select(oRange.getCellByPosition(0,rowNb))		
		oRange.getCellByPosition(0,rowNb).setString(result)
		
	next rowNb
end Sub









