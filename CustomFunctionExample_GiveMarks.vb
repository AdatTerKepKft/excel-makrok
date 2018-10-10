Function giveMarks(rng As Range, maxPoints As Integer)
'A makró célja, hogy létrehozzon egy olyan függvényt, amely segítségével az összpontszám'
'kijelölésével, valamint a maximális pontszám megadásával meghatározható az érdemjegy.'

'Ellenőrizzük, hogy valóban csak egy cellát jelölt-e ki a függvény hívása során.
If rng.Cells.Count > 1 Then
 'Ha nem, hibaüzenetet írunk ki és kilépünk.'
 giveMarks = "Csak egy cellát jelölj ki!"
 Exit Function
End If

'A megszerzett pontszám (rng.Value) és a maximális pontszám (maxPoints) alapján százalékot'
'számolunk, amely alapján meghatározzuk az érdemjegyet.'
If (rng.Value / maxPoints * 100) <= 50 Then
 giveMarks = 1
ElseIf (rng.Value / maxPoints * 100) < 60 Then
 giveMarks = 2
ElseIf (rng.Value / maxPoints * 100) < 70 Then
 giveMarks = 3
ElseIf (rng.Value / maxPoints * 100) < 85 Then
 giveMarks = 4
Else
 giveMarks = 5
End If

End Function