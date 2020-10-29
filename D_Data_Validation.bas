Attribute VB_Name = "D_Data_Validation"
'F�ggv�ny ami l�trehozza a leg�rd�l�men�t (a dashboard negyedik oszlop�ba) a megfelel� sheet megfelel� oszlop�nak egyes �rt�keivel
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

'Addig minuszolja loopban a rangeendet am�g m�r nem �res az utols� cella
Function rangeEndf(ByRef rangeBegin As String, rangeEndn As Integer) As String

Application.EnableEvents = False

Do While IsEmpty(range(CStr("" & rangeBegin & rangeEndn & "")).Value)
        rangeEndn = rangeEndn - 1
        rangeEndf = rangeBegin & rangeEndn
        Loop

Application.EnableEvents = True

End Function
