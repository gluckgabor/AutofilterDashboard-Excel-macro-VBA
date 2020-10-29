VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Sub Workbook_Open()
Set AnyWorkbook = New clsApplicationEvents
Set AnyWorkbook.XL = Excel.Application
End Sub

Public Sub AutofilterDashboard()

ActiveSheet.Buttons.Delete

Application.EnableEvents = False

Dim Header As String
Dim rangeBegin As String, rangeBeginHidden As String
Dim rangeEnd As String
Dim Rng As String
Dim firstEmptyRowNumber As String
Dim rangeEndfv As String

Dim rangeEndn As Integer
Dim a As Integer
Dim utolsoirtdashboardsor As Integer
Dim emptyRowNum As Integer
Dim currentregionRowcount As Integer
Dim currentregionColumncount As Integer

Dim r As range, b As range
Dim filterDropDownHeader As range
 
    
    '(0)Teardown action, just to be sure
    currentregionRowcount = range("A1").currentregion.Rows.Count
    currentregionColumncount = range("A1").currentregion.Columns.Count
    
    range(Cells(currentregionRowcount + 1, 1), Cells(currentregionRowcount * 2, currentregionColumncount)).Select
    range(Cells(currentregionRowcount + 1, 1), Cells(currentregionRowcount * 2, currentregionColumncount)).ClearContents
 
 
    '(1)Datavalidation-értékkészlet-létrehozó ciklus
              
        'itt kezdõdik az átalakított lista
    emptyRowNum = firstEmptyRowNum(1) + 2
    
    Set r = range(Cells(1, 1), Cells(1, currentregionColumncount)) 'sor, oszlop
    
    For Each b In r 'a loop
        
        If b.Value = "X" Or b.Value = "x" Then 'Check upper row whether there is X. If nth cell contains X generate 4 cells for it below. if there is:
            
            Debug.Print "b.Address: " & b.Address
                'csak a sorokkal kalkulál, viszont, mivel az elsõ oszlopba beírtuk, hogy melyik oszlopra vonatkozik az adott dashboard sor ezért bármelyik
                'sor alapjá tudja, hogy melyik oszloppal foglalkozzon
                         
            Call PopulateDV(b, emptyRowNum)
    
        End If
        
    Next 'if not then just get back to beginning of loop till all cells are checked in row




    '(2)Dashboardfejléc létrehozása
    a = emptyRowNum + emptyRowNum - 2 * 2
    
    Sheets(Sheet1.Name).range(Cells(a, 1), Cells(a, 1)).Value = "Columnletter"
    Sheets(Sheet1.Name).range(Cells(a, 2), Cells(a, 2)).Value = "Column name"
    Sheets(Sheet1.Name).range(Cells(a, 3), Cells(a, 3)).Value = "Current filtering"
        
Set filterDropDownHeader = Sheets(Sheet1.Name).range(Cells(a, 4), Cells(a, 4))
    filterDropDownHeader.Value = "Desired filtering"


    
    
    '(3)Dashboard-rekord-létrehozó ciklus
    utolsoirtdashboardsor = 0
    For Each b In r 'a loop
        
        If b.Value = "X" Or b.Value = "x" Then 'Check upper row whether there is X. If nth cell contains X generate 4 cells for it below. if there is:
            
            Header = b.Offset(1, 0).Value  'Szerezze meg az adott oszlop fejlécnevét, x-hez képest egy sorral lejjebb
            
            rangeEndn = firstEmptyRowNum(b.Column) 'Szerezze meg az adott oszlop tartományát, ehhez kell az adott oszlophoz tartozó _
            firstEmptyRowNumber függvény meghívása az adott oszlopra, melynek bemeneti értéke az oszlop sorszáma
                        
            rangeBegin = Mid(b.Address, 2, InStr(2, b.Address, "$") - 2) 'b.Address megadja a betûjelet
            rangeEndfv = rangeEndf(rangeBegin, rangeEndn)
            Rng = "D" & (a + 1 + utolsoirtdashboardsor)
             
            'Insert to 1nd cell
            range(Cells(a + 1 + utolsoirtdashboardsor, 1), Cells(a + 1 + utolsoirtdashboardsor, 1)).Value = rangeBegin
            
            'Insert to 2nd cell Header value to right place
            range(Cells(a + 1 + utolsoirtdashboardsor, 2), Cells(a + 1 + utolsoirtdashboardsor, 2)).Value = Header
            
            'Insert to 3nd cell
            range(Cells(a + 1 + utolsoirtdashboardsor, 3), Cells(a + 1 + utolsoirtdashboardsor, 3)).Value = "" & "=AutofilterDashboard6.xlsm!DispCriteria(" & rangeBegin & ":" & rangeBegin & ")" & ""
        
         utolsoirtdashboardsor = utolsoirtdashboardsor + 1
         
        End If
    
    Next 'if not then just get back to beginning of loop till all cells are checked in row

        
        
        
    '(3)Dashboard utolsó cellában dropdownlistet frissítõ ciklus
    'BOM: ciklikusan végig kell menni a bal szélsõ range-n, betûjel kell és offset 3-al jobbra az Rng
    
    Dim SplitRow As Integer
    SplitRow = CStr(Sheets(Sheet1.Name).range("A:A").Find("Columnletter").Offset(-2, 0).Row) + 2
    
    For Each b In r
    
        If b.Value = "X" Or b.Value = "x" Then
            
            Dim ColumnIDRange As range
            Set ColumnIDRange = Sheets(Sheet1.Name).range(Cells(a + 1, 1), Cells(Sheets(Sheet1.Name).range("A65536").End(xlUp).Cells.Row, 1))
            Debug.Print ColumnIDRange.Address
            
            Dim locationOfCID As range
            Set locationOfCID = ColumnIDRange.Find(CStr(Mid(b.Address, 2, InStr(2, b.Address, "$") - 2)))
                    
            'a dropdownlist insert-hez használandó cella
                    
            'címének kiderítése: offset jobbra 3 oszloppal levõ cella címe kell, mint string
            'Rng = CStr(locationOfCID.Offset(0, 3).Address)
            
            rangeBegin = CStr(Mid(b.Address, 2, InStr(2, b.Address, "$") - 2)) 'b.Address megadja a betûjelet
            rangeEndn = CStr(Sheets(Sheet1.Name).range("A:A").Find("Columnletter").Offset(-2, 0).Row) 'Szerezze meg az adott oszlop tartományát, ehhez kell az adott oszlophoz tartozó
            
            'firstEmptyRowNumber függvény meghívása az adott oszlopra, melynek bemeneti értéke az oszlop sorszáma
                       
            
            'Addig minuszolja loopban a rangeendet amíg már nem üres az utolsó cella
            'Function rangeEndf(ByRef rangeBegin As String, rangeEndn As Integer) As String
            rangeEndfv = rangeEndf(rangeBegin, rangeEndn)
                          
            Rng = "D" & (a + 1)
            
            'ByVal rangeBegin As String, rangeEndfv As String, ByVal Rng As String
            Call RefreshDV(emptyRowNum, rangeBegin, rangeEndfv, Rng)
            
            a = a + 1
                        
        End If
        
    Next

Call OrganizeWindow(SplitRow)

Application.EnableEvents = True
  
End Sub
