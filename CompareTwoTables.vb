Sub CompareTwoTables()

'A makró célja, hogy két adattábla tartalmát összehasonlítsa, esetleges eltéréseiket jelezze és ezekről listát készítsen.
'A makró igényel némi előkészítést.
'A két adattábla adatait egy munkafüzet egy-egy munkalapjára kell másolni, a régebbit "OLD", az újabbat "NEW" névvel kell
'ellátni (idézőjelek nélkül).

'Strukturális különbségek

'Ellenőrizzük, a két munkalapon azonos számú sor szerepel-e
old_rows_count = Worksheets("OLD").UsedRange.Rows.Count
new_rows_count = Worksheets("NEW").UsedRange.Rows.Count
If old_rows_count <> new_rows_count Then
    MsgBox ("A sorok száma eltér a két munkalapon! OLD: " & old_rows_count & ", NEW: " & new_rows_count)
    Exit Sub
End If

'Ellenőrizzük, a két munkalapon azonos számú oszlop szerepel-e
old_cols_count = Worksheets("OLD").UsedRange.Columns.Count
new_cols_count = Worksheets("NEW").UsedRange.Columns.Count
If old_cols_count <> new_cols_count Then
    MsgBox ("Az oszlopok száma eltér a két munkalapon! OLD: " & old_cols_count & ", NEW: " & new_cols_count)
    Exit Sub
End If

'Ellenőrizzük, a két munkalapon megegyezik-e a fejléc (az első sor értékei, a sorrendet is figyelembe véve)
For c = 1 To old_cols_count
    If Worksheets("OLD").Cells(1, c).Value <> Worksheets("NEW").Cells(1, c).Value Then
        MsgBox ("A fejléc eltér a " & c & ". oszlopban! OLD: " & Worksheets("OLD").Cells(1, c).Value & ", NEW: " & Worksheets("NEW").Cells(1, c).Value)
        Exit Sub    
    End If
Next c

'Ellenőrizzük, a két munkalapon megegyeznek-e a sor azonosítók (az első oszlop értékei, a sorrendet is figyelembe véve)
For r = 1 To old_rows_count
    If Worksheets("OLD").Cells(r, 1).Value <> Worksheets("NEW").Cells(r, 1).Value Then
        MsgBox ("A sorazonosító eltér a " & r & ". sorban! OLD: " & Worksheets("OLD").Cells(r, 1).Value & ", NEW: " & Worksheets("NEW").Cells(r, 1).Value)
        Exit Sub    
    End If
Next r

'Tartalmi különbségek

differences = 0
Worksheets.Add().Name = "DIFF"
Worksheets("DIFF").Cells(1,1).Value = "ROW"
Worksheets("DIFF").Cells(1,2).Value = "COL"
Worksheets("DIFF").Cells(1,3).Value = "OLD_VALUE"
Worksheets("DIFF").Cells(1,4).Value = "NEW_VALUE"

For r = 2 To old_rows_count
    For c = 2 To old_cols_count
        If Worksheets("OLD").Cells(r,c).Value <> Worksheets("NEW").Cells(r,c).Value Then
            Worksheets("DIFF").Cells(differences + 2,1).Value = r
            Worksheets("DIFF").Cells(differences + 2,2).Value = c
            'Ha a sor- és oszlopazonosítót szeretnéd a sor és oszlop száma helyett kiíratni, a fenti két sort cseréld ki a lenti két sorra
            'Worksheets("DIFF").Cells(differences + 2,1).Value = Worksheets("OLD").Cells(r, 1).Value
            'Worksheets("DIFF").Cells(differences + 2,2).Value = Worksheets("OLD").Cells(1, c).Value
            Worksheets("DIFF").Cells(differences + 2,3).Value = Worksheets("OLD").Cells(r,c).Value
            Worksheets("DIFF").Cells(differences + 2,4).Value = Worksheets("NEW").Cells(r,c).Value
            differences = differences + 1
        End If
    Next c
Next r

If differences = 0 Then
    Application.DisplayAlerts = False
    Worksheets("DIFF").Delete
    Application.DisplayAlerts = True
    MsgBox ("A két adattábla teljesen megegyezik!")
Else
    MsgBox ("A két adattábla strukturálisan megegyezik, tartalmi különbségeik (" & differences & " darab) a DIFF munkalapon láthatók!")
End If

End Sub
