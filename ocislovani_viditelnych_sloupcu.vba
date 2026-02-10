Sub OcislovatViditelneSloupce()
    Dim oblast As Range
    Dim sloupec As Range
    Dim i As Integer
    
    ' Pracujeme s tím, co jste zrovna označili myší
    Set oblast = Selection
    i = 1
    
    ' Projdeme každý sloupec v označené oblasti
    For Each sloupec In oblast.Columns
        ' Pokud není sloupec skrytý...
        If sloupec.EntireColumn.Hidden = False Then
            ' ...napíšeme do první buňky označeného sloupce číslo
            sloupec.Cells(1, 1).Value = i
            i = i + 1
        End If
    Next sloupec
    
    MsgBox "Hotovo! Očíslováno " & (i - 1) & " viditelných sloupců.", vbInformation
End Sub
