Attribute VB_Name = "F1_ClearAllFilters_Btn1_Click"
Option Private Module
Public Sub Btn1_Click()

Dim locationOfDFtext As range
Dim locationOflastfilterCell As Variant
Dim szuroOszlopRange As range
    
            Set locationOfDFtext = Sheets(Sheet1.Name).range("D:D").Find("Desired filtering").Offset(1, 0)
            
            Set locationOflastfilterCell = Sheets(Sheet1.Name).range("B65536").End(xlUp).Offset(0, 2)
            
            Set szuroOszlopRange = Sheets(Sheet1.Name).range("" & locationOfDFtext.Address & ":" _
            & locationOflastfilterCell.Address & "")
            
            Debug.Print locationOfDFtext.Address
            Debug.Print locationOflastfilterCell.Address
            Debug.Print szuroOszlopRange.Address
                
            szuroOszlopRange.ClearContents
            
        a = firstEmptyRowNum(1) + 2
        
        Call Btn_Click
        
        Dim SplitRow As Integer
    SplitRow = CStr(Sheets(Sheet1.Name).range("A:A").Find("Columnletter").Offset(-2, 0).Row) + 2
    
        Call OrganizeWindow(SplitRow)

End Sub
