Sub ListFileNamesInFolder()
'A makró célja, hogy kilistázza egy munkalapra az egy mappában található fájlokat.'

'Változók definiálása.'
Dim path As String
Dim file As String
Dim row As Long

'A fájlokat tartalmazó mappa elérési útjának bekérése'
path = Application.InputBox("Adja meg a mappa elérési útját: ", "Listázandó mappa elérési útja", Type:=2)

'Új munkalap beszúrása, ami a fájlok listáját fogja tartalmazni.'
Worksheets.Add Before := Worksheets(1)

'Az imént létrehozott munkalap átnevezése'
Worksheets(1).Name = "LISTA"

'Fejléc kiírása'
Worksheets(1).Cells(1,1).Value = "Fájlok"

'Melyik sorba kerüljön az első listázott fájl?'
row = 2

'Első fájlnév felolvasása'
file = Dir(path)

'A ciklus kezdete. A ciklusmag (A While... és Wend közötti sorok) addig fog újra és újra lefutni,'
'míg a file változó rendelkezik értékkel (pontosabban: nem üres szöveg az értéke)'
While file <> ""
    'A fájl nevének kiírása a következő üres sorba.'
    Worksheets(1).Cells(row, 1).Value = file
    'A következő fájl felolvasása'
    file = Dir()
    'Az eredmény munkalapon a sor számának növelése'
    row = row + 1
Wend

End Sub
