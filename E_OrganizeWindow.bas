Attribute VB_Name = "E_OrganizeWindow"
Sub OrganizeWindow(ByVal SplitRow As Integer)

Application.EnableEvents = False

Sheets(Sheet1.Name).Activate
 
 ActiveWindow.WindowState = xlMaximized
 ActiveWindow.Zoom = 100
 ActiveWindow.SplitRow = 0 'Középen vízszintesen kettéosztja a képmezõt
 ActiveWindow.SplitRow = 16 'Középen vízszintesen kettéosztja a képmezõt

 ActiveWindow.Panes(1).Activate
 ActiveWindow.Panes(1).ScrollRow = 1 'A felsõ képmezõben a táblázat tetejére ugrik

 ActiveWindow.Panes(2).Activate
 ActiveWindow.Panes(2).ScrollRow = SplitRow 'Az alsó képmezõben a dashboard tetejére ugrik
 range(Cells(SplitRow, 1), Cells(SplitRow, 1)).Select
 
 Application.EnableEvents = True
 
  With ActiveSheet
                    
                    .EnableCalculation = False
                    .EnableCalculation = True
                    .Calculate
                        
  End With
 
End Sub
