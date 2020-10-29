VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsApplicationEvents"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents oApp As Excel.Application
Attribute oApp.VB_VarHelpID = -1

Property Set XL(Application As Excel.Application)
Set oApp = Application
End Property
Property Get XL() As Excel.Application
Set XL = oApp
End Property

Private Sub Class_Initialize()

End Sub

'A dashboard harmadik oszlopában kiválasztott érték alapján vezérli az autofiltert
Public Sub oApp_SheetChange(ByVal Sh As Object, ByVal Target As range)

Application.EnableEvents = False
     
            Call OrganizeWindow(a)
            
                With ActiveSheet
                    
                    .EnableCalculation = False
                    .EnableCalculation = True
                    .Calculate
                        
                End With
                
         End If
        
    End With
    
Application.EnableEvents = True
    
End Sub
