Attribute VB_Name = "C_PopulateDV"
Sub PopulateDV(ByVal b As range, ByVal emptyRowNum As Integer)
  
  Application.EnableEvents = False
  
  Dim actuallyWorkedColumnNo As Integer
  actuallyWorkedColumnNo = b.Column
  
  With ActiveSheet
                      
    Application.EnableEvents = False
    .EnableCalculation = False
    
  
    With Sheets(Sheet1.Name)
    
        Application.EnableEvents = False
    
    '(1)az �rt�kk�szetet eleve lentre m�solja �s rendezi, a datavalidation erre hivatkozik mindig
    '(4)lent aktualiz�l�dik a fenti lista alapj�n az �trendezett lista, ett�l automatikusan friss�l a data validation is
                
            'beikszelt tartom�ny m�sol�sa 'majd beilleszt�k az �j hely�re a dashboard alatt
            Dim copiedRangeRowCount As Integer
            copiedRangeRowCount = range("" & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & "2" & ":" & _
            Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum - 4 & "").Offset(1, 0).Rows.Count
            
            Dim columnLetterID As String
            columnLetterID = Mid(b.Address, 2, InStr(2, b.Address, "$") - 2)
            
            range("" & columnLetterID & "3" & ":" & _
            columnLetterID & emptyRowNum - 3 & "").Offset(0, 0).Copy _
            range("" & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum)
            
            'a k�v. sor az�rt kell, hogy ne hagyjon szaggatottvonal� kijel�l�st a copy m�velet
            Application.CutCopyMode = False
                                
            'ott a tartom�ny aktualiz�l�sa, mentes�t�se (duplicates)
            Debug.Print "mentes�tett tartom�ny: " & "" & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum & ":" _
                            & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum + copiedRangeRowCount - 1 & ""
            
            range("" & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum & ":" _
                            & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum + copiedRangeRowCount - 1 & "").Select
                            
            range("" & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum & ":" _
                            & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum + copiedRangeRowCount - 1 & ""). _
                            RemoveDuplicates Columns:=1, Header:=xlNo
                     
            'Range containing blank cells
            Dim rsltRng As range
            Set rsltRng = range("" & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum & ":" _
                            & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) & emptyRowNum + 1 + copiedRangeRowCount & "")
            Debug.Print "rsltRng.Address: " & rsltRng.Address
            
            'Loop through rsltRng and systematically move upwards cell by cell if a certain cell contains data
            Dim cell As range
            For Each cell In rsltRng
            
                Application.EnableEvents = False
                
                If IsEmpty(cell) Then
                    
                    cell.Offset(1, 0).Cut cell
                    
                End If
                        
            Next
                        
                        
        'oszlop
        'b.Address megadja a bet�jelet
        Dim rangeBegin As String
        rangeBegin = Mid(b.Address, 2, InStr(2, b.Address, "$") - 2)
        
        'h�ny sorr�l van sz� eredetileg: copiedRangeRowCount
        Dim rangeEndn As Integer
        rangeEndn = emptyRowNum + 1 + copiedRangeRowCount
        
        'User defined function (UDF), ami addig minuszolja loopban a rangeendet am�g m�r nem �res az utols� cella
        'stringet ad vissza eredm�ny�l
        Dim rangeEndfv As String
        rangeEndfv = rangeEndf(rangeBegin, rangeEndn)
                
        Debug.Print "oszlop bet�jele: " & Mid(b.Address, 2, InStr(2, b.Address, "$") - 2)
              
    End With
    
        Application.EnableEvents = False
        
    .EnableCalculation = True
    .Calculate
    
  End With
  
  Application.EnableEvents = True

End Sub
