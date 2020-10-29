Attribute VB_Name = "F_ButtonAdder"
Option Private Module
Public Sub ButtonAdder()
    Dim btn As Button
    Dim t As range
    Dim t1 As range
    Dim locationOfDFtext As range
  
        Application.ScreenUpdating = False
        

  With Sheets(Sheet1.Name).range("D:D")
     
      Set locationOfDFtext = .Find("Desired filtering")
      Debug.Print "locationOfDFtext.address: " & locationOfDFtext.Offset(1, 1).Address
      Set t = locationOfDFtext.Offset(1, 1)
      Set t1 = locationOfDFtext.Offset(2, 1)
        
        Set btn = ActiveSheet.Buttons.Add(t.Left, t.Top, t.Width, t.Height)
        With btn
          .OnAction = "AutofilterDashboard6.xlsm!Btn_Click"
          .Caption = "  Refresh all"
          .Name = "Btn"
        End With
        
        Set btn = ActiveSheet.Buttons.Add(t1.Left, t1.Top, t1.Width, t1.Height)
        With btn
          .OnAction = "AutofilterDashboard6.xlsm!Btn1_Click"
          .Caption = "  Clear all filters"
          .Name = "Btn1"
        End With
      
    End With
  Application.ScreenUpdating = True
End Sub
