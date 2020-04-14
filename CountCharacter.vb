Function CountCharacter(txt As String, char As String, Optional cs As Boolean = True) As Long
    If cs Then
        CountCharacter = (Len(txt) - Len(Replace(txt, char, ""))) / Len(char)
    Else
        CountCharacter = (Len(txt) - Len(Replace(LCase(txt), LCase(char), ""))) / Len(char)
    End If
End Function