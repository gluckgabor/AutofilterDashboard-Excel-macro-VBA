Attribute VB_Name = "C_RefreshDV"
Sub RefreshDV(ByVal emptyRowNum As Integer, ByVal rangeBegin As String, rangeEndfv As String, ByVal Rng As String)

        'Insert to 4th cell the datavalidation dropdown !!!!!!!!!!!!!!!!!!!!!!!!!!!
        'Sheets(Sheet1.Name).range(Cells(a + 1, 4), Cells(a + 1, 4)).Value = Data_Validation(rangeBegin, rangeEndfv, Rng)
        range(Rng).Value = Data_Validation(emptyRowNum, rangeBegin, rangeEndfv, Rng)
        
        Call F_ButtonAdder.ButtonAdder
        
End Sub
