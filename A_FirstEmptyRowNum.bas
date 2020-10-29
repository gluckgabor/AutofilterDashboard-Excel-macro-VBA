Attribute VB_Name = "A_firstEmptyRowNum"
'Megmondja, hogy mi az elsõ üres sor száma a filterezést nem figyelembevéve
Function firstEmptyRowNum(ByVal column_with_data As Integer) As Integer

Application.EnableEvents = False

If Sheets(Sheet1.Name).AutoFilterMode = False Then 'ha nincs szûrés
    firstEmptyRowNum = Sheets(Sheet1.Name).SpecialCells(xlCellTypeLastCell).Offset(1, 0).Row
    
Else
    With Sheets(Sheet1.Name).AutoFilter.range 'ha van autofilter (ha végig select all-os az is számít)
            firstEmptyRowNum = .Row + .Rows.Count
    End With
End If

Application.EnableEvents = True

End Function
'Megmondja, hogy mi az elsõ üres sor száma az adott oszlopban
Function firstEmptyRowNumC(ByVal column_with_data As Long) As Long

Application.EnableEvents = False

If ActiveSheet.AutoFilterMode = False Then
    firstEmptyRowNumC = range("B65536").End(xlUp).Offset(1, 0).Row
Else
    With ActiveSheet.AutoFilter.range(firstEmptyRowNum(column_with_data), column_with_data)

       firstEmptyRowNumC = .Cells(.Rows.Count, column_with_data).End(xlUp).Row
    End With
End If

Application.EnableEvents = True

End Function




