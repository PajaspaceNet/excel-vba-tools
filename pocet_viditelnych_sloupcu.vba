Sub PocetViditelnychSloupcu()
    Dim radek As Range
    ' Vezmeme první řádek z vašeho výběru a spočítáme v něm viditelné buňky
    Set radek = Selection.Rows(1).SpecialCells(xlCellTypeVisible)
    
    MsgBox "Počet viditelných sloupců ve výběru: " & radek.Count, vbInformation, "Počítadlo"
End Sub
