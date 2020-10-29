Attribute VB_Name = "D_Data_Validation"
'Függvény ami létrehozza a legördülõmenüt (a dashboard negyedik oszlopába) a megfelelõ sheet megfelelõ oszlopának egyes értékeivel
Function Data_Validation(ByVal emptyRowNum As Integer, rangeBegin As String, rangeEndfv As String, Rng As String) As String

Application.EnableEvents = False
    
    Sheets(Sheet1.Name).range("" & Rng & "").Select
    With Selection.Validation
        .Delete
        .Add Type:=xlValidateList, AlertStyle:=xlValidAlertStop, Operator:=xlBetween, Formula1:="" & "=" & rangeBegin & emptyRowNum & ":" & rangeEndfv & ""
        .IgnoreBlank = True
        .InCellDropdown = True
        .InputTitle = ""
        .ErrorTitle = ""
        .InputMessage = ""
        .ErrorMessage = ""
        .ShowInput = True
        .ShowError = True
    End With
    
    Application.EnableEvents = True

End Function

'Addig minuszolja loopban a rangeendet amíg már nem üres az utolsó cella
Function rangeEndf(ByRef rangeBegin As String, rangeEndn As Integer) As String

Application.EnableEvents = False

Do While IsEmpty(range(CStr("" & rangeBegin & rangeEndn & "")).Value)
        rangeEndn = rangeEndn - 1
        rangeEndf = rangeBegin & rangeEndn
        Loop

Application.EnableEvents = True

End Function
