Attribute VB_Name = "B_DispCriteria"
'A második oszlopba kitölti, hogy milyen szûréssel érkezett a fájl
Function DispCriteria(Rng As range) As String

Application.EnableEvents = False

    Dim Filter As String

    Filter = ""
    On Error GoTo done
    With Rng.Parent.AutoFilter
        If Intersect(Rng, .range) Is Nothing Then GoTo done
        With .Filters(Rng.Column - .range.Column + 1)
            If Not .On Then GoTo done
            Filter = .Criteria1
            Select Case .Operator
                Case xlAnd
                    Filter = Filter & " AND " & .Criteria2
                Case xlOr
                    Filter = Filter & " OR " & .Criteria2
            End Select
        End With
    End With
done:
    DispCriteria = Filter

Application.EnableEvents = True

End Function
