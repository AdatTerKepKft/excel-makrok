Sub CreateAndLinkTableOfContent()
'A makró egy tetszőleges nevű munkalapot szúr be a meglévőek elé. Erre a munkalapra egy tartalomjegyzéket készít a többi munkalapot listázva, hivatkozást is elhelyezve az egyes munkalapokra.'

'Változók definiálása'
Dim sheetName, backLinkText As String
Dim sheetCount, tocRow As Long
Dim backLink As Integer

'Kérdezze meg a felhasználótól, mi legyen a tartalomjegyzék munkalapjának a neve!'
sheetName = InputBox("Mi legyen a tartalomjegyzék munkalapjának neve?", "Tartalomjegyzék munkalapjának neve")

'Kérdezze meg, szeretnénk-e vissza gombot elhelyezni a munkalapokon?'
backLink = MsgBox("Legyen-e egy Vissza logikájú link a munkalapok első sorában?", 4, "Vissza logikájú link")
'Ha igen, kérdezze meg, mi legyen a szöveg?'
If backLink = 6 Then
	backLinkText = InputBox("Mi legyen a Vissza logikájú link felirata?", "Vissza logikájú link felirata")
End If

'Szúrjon be egy új munkalapot a meglévők elé a legelső helyre.'
ActiveWorkbook.Sheets.Add Before:=Worksheets(1)
'Adja az új munkalapnak a felhasználó által megadott nevet!'
Worksheets(1).Name = sheetName

'Gyűjtse össze sorban, milyen munkalapok vannak a munkafüzetben...'
sheetCount = ActiveWorkbook.Sheets.Count

For tocRow = 1 To sheetCount
	
	'... és írja ezeket a Tartalomjegyzék munkalapra sorszámokkal ellátva.'
	Worksheets(1).Cells(tocRow, 1).Value = tocRow
	Worksheets(1).Cells(tocRow, 2).Value = Worksheets(tocRow).Name
	
	'A munkalapok neveit tegye linkké, hogy a tartalomjegyzékben rákattintva a névre a megfelelő munkalapra ugorjon.'
	With Worksheets(1)
		.Hyperlinks.Add Anchor:=.Cells(tocRow,2), _
		Address:="", _
		SubAddress:="'" & Worksheets(tocRow).Name & "'!A1", _
		TextToDisplay:=Worksheets(tocRow).Name
	End With

	'Haladjon végig a munkalapokon és minden munkalapra szúrjon be egy sort az összes többi sor elé.
	If backLink = 6 And tocRow > 1 Then
		Worksheets(tocRow).Rows(1).EntireRow.Insert
		
		'Az új sorokba (az A1 cellába) írja bele a megadott szöveget, amit tegyen linkké, ami a tartalomjegyzékre mutat, hogy így vissza lehessen navigálni a tartalomjegyzék munkalapjára.'
		Worksheets(tocRow).Cells(1,1).Value = backLinkText
		With Worksheets(tocRow)
			.Hyperlinks.Add Anchor:=.Cells(1,1), _
			Address:="", _
			SubAddress:="'" & sheetName & "'!A1", _
			TextToDisplay:=backLinkText	
		End With
	End If
	
Next tocRow

End Sub