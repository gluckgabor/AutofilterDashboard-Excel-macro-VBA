Attribute VB_Name = "E_OrganizeWindow"
Sub OrganizeWindow(ByVal SplitRow As Integer)

Application.EnableEvents = False

Sheets(Sheet1.Name).Activate
 
 ActiveWindow.WindowState = xlMaximized
 ActiveWindow.Zoom = 100
 ActiveWindow.SplitRow = 0 'K�z�pen v�zszintesen kett�osztja a k�pmez�t
 ActiveWindow.SplitRow = 16 'K�z�pen v�zszintesen kett�osztja a k�pmez�t

 ActiveWindow.Panes(1).Activate
 ActiveWindow.Panes(1).ScrollRow = 1 'A fels� k�pmez�ben a t�bl�zat tetej�re ugrik

 ActiveWindow.Panes(2).Activate
 ActiveWindow.Panes(2).ScrollRow = SplitRow 'Az als� k�pmez�ben a dashboard tetej�re ugrik
 range(Cells(SplitRow, 1), Cells(SplitRow, 1)).Select
 
 Application.EnableEvents = True
 
  With ActiveSheet
                    
                    .EnableCalculation = False
                    .EnableCalculation = True
                    .Calculate
                        
  End With
 
End Sub
