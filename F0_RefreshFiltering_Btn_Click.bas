Attribute VB_Name = "F0_RefreshFiltering_Btn_Click"
Option Private Module
Public Sub Btn_Click()

'Változók deklarálása
Dim a As Integer
Dim rangeEndn As Integer
Dim rngInt As Integer
Dim rngOsszesitoInt As Integer
Dim i As Integer

Dim dashFourthColumnRange As range
Dim cell As range
Dim targetValue As range
Dim targetAddressRng As range
Dim hiddenSheetRange As range
Dim r As range
Dim b As range
Dim rsltRngBlanks As range

Dim col As Variant

Dim rangeBegin As String
Dim rangeEndfv As String
Dim firstcell As String, lastCell As String, rangeEnd As String, Rng As String, firstEmptyRowNumber As String

Dim SheetOne As Worksheet
Dim hiddenWorksheet As Worksheet

Dim rsltRng As Object

Dim letezik As Boolean

Dim szuroOszlopRange As range
Dim szuroCellaRange As range
   
    'Check filtered range
    a = firstEmptyRowNum(1) + 2
           
    firstcell = range(Cells(a, 4), Cells(a, 4)).Offset(1, 0).Address
    lastCell = range("B65536").End(xlUp).Offset(1, 2).Address
    
    Set dashFourthColumnRange = Sheets(Sheet1.Name).range("" & firstcell & ":" & lastCell & "")
            
    With Application.ActiveWorkbook.ActiveSheet
                    
        'a D oszlop "desired filtering" range-n végigszaladunk és ahol van beállítva kívánt szûrõfeltétel, annál a cellánál
        'elkönyvelünk egy egyest(amúgy nullát), ezt egy gyûjtõ-változóba beiratni, ebbõl tudja a program, hogy esetleg
        'mindent kell-e mutatnia az egész táblázaton, vagy van beállított szûrõfeltétel
        rngOsszesitoInt = 0
            
        For Each b In dashFourthColumnRange
            
            If IsEmpty(b) Then
            rngInt = 0
            Else: rngInt = 1
            End If
               
            rngOsszesitoInt = rngOsszesitoInt + rngInt
        Next
         
         
         
        'megvizsgálni az összesítõ változót:
        
        '(1)ha üres akkor show all,
        If rngOsszesitoInt = 0 Then
            ActiveSheet.AutoFilter.ShowAllData
                   
        '(2)ha nem üres, azaz beírtunk a dashboardon szûrõértéket az alapján szûri fent a megfelelõ oszlopot
        Else
            'Egy loop végigmegy a datavalidation oszlopon és a beállított értékeknek megfelelõen rendezi a fenti táblát
        
            With Sheets(Sheet1.Name).range("D:D")
        
                Dim locationOfDFtext As range
                Set locationOfDFtext = Sheets(Sheet1.Name).range("B:B").Find("Column name").Offset(2, 2)
                Debug.Print "locationOfDFtext.address: " & locationOfDFtext.Address
                                            
                Set t = locationOfDFtext
                Debug.Print "t.Address: " & t.Address

                Debug.Print "Dash utolsó sor D oszlop metszete címe " & Sheets(Sheet1.Name).range("B65536").End(xlUp).Cells.Offset(0, 2).Address
                
                'utolsó kitöltött cella keresése
                Set szuroOszlopRange = _
                range("" & t.Address & _
                ":" & Sheets(Sheet1.Name).range("B65536").End(xlUp).Cells.Offset(0, 2).Address & "")
                
                Debug.Print "szuroOszlopRange.Address: "; szuroOszlopRange.Address
                
                
                    For Each szuroCellaRange In szuroOszlopRange
                    
                    Debug.Print "szuroCellaRange.Address: "; szuroCellaRange.Address
                    
                        'hogy frissíthessük a filterezést, azonosítjuk a frissítendõ oszlopot
                        Set targetAddressRng = szuroCellaRange.Offset(0, -3)
                        
                        If IsEmpty(szuroCellaRange) Then
                        
                            With Sheets(Sheet1.Name)
                            
                            'itt történik meg a lényeg, a szûrés !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                            .range("" & targetAddressRng & "2" & ":" & targetAddressRng & "2" & "").AutoFilter Field:=range(targetAddressRng & 1).Column

                            End With
                            
                        Else
                        
                            With Sheets(Sheet1.Name)
                            
                            'itt történik meg a lényeg, a szûrés !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                            .range("" & targetAddressRng & "2" & ":" & targetAddressRng & "2" & "").AutoFilter Field:=range(targetAddressRng & 1).Column, Criteria1:=szuroCellaRange

                            End With
                        
                        End If
                        
                    Next
                    
            End With
        
        End If
               
    End With
    
    Dim SplitRow As Integer
    SplitRow = CStr(Sheets(Sheet1.Name).range("A:A").Find("Columnletter").Offset(-2, 0).Row) + 2
    
Call OrganizeWindow(SplitRow)
    
End Sub
