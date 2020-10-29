Attribute VB_Name = "F0_RefreshFiltering_Btn_Click"
Option Private Module
Public Sub Btn_Click()

'V�ltoz�k deklar�l�sa
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
                    
        'a D oszlop "desired filtering" range-n v�gigszaladunk �s ahol van be�ll�tva k�v�nt sz�r�felt�tel, ann�l a cell�n�l
        'elk�nyvel�nk egy egyest(am�gy null�t), ezt egy gy�jt�-v�ltoz�ba beiratni, ebb�l tudja a program, hogy esetleg
        'mindent kell-e mutatnia az eg�sz t�bl�zaton, vagy van be�ll�tott sz�r�felt�tel
        rngOsszesitoInt = 0
            
        For Each b In dashFourthColumnRange
            
            If IsEmpty(b) Then
            rngInt = 0
            Else: rngInt = 1
            End If
               
            rngOsszesitoInt = rngOsszesitoInt + rngInt
        Next
         
         
         
        'megvizsg�lni az �sszes�t� v�ltoz�t:
        
        '(1)ha �res akkor show all,
        If rngOsszesitoInt = 0 Then
            ActiveSheet.AutoFilter.ShowAllData
                   
        '(2)ha nem �res, azaz be�rtunk a dashboardon sz�r��rt�ket az alapj�n sz�ri fent a megfelel� oszlopot
        Else
            'Egy loop v�gigmegy a datavalidation oszlopon �s a be�ll�tott �rt�keknek megfelel�en rendezi a fenti t�bl�t
        
            With Sheets(Sheet1.Name).range("D:D")
        
                Dim locationOfDFtext As range
                Set locationOfDFtext = Sheets(Sheet1.Name).range("B:B").Find("Column name").Offset(2, 2)
                Debug.Print "locationOfDFtext.address: " & locationOfDFtext.Address
                                            
                Set t = locationOfDFtext
                Debug.Print "t.Address: " & t.Address

                Debug.Print "Dash utols� sor D oszlop metszete c�me " & Sheets(Sheet1.Name).range("B65536").End(xlUp).Cells.Offset(0, 2).Address
                
                'utols� kit�lt�tt cella keres�se
                Set szuroOszlopRange = _
                range("" & t.Address & _
                ":" & Sheets(Sheet1.Name).range("B65536").End(xlUp).Cells.Offset(0, 2).Address & "")
                
                Debug.Print "szuroOszlopRange.Address: "; szuroOszlopRange.Address
                
                
                    For Each szuroCellaRange In szuroOszlopRange
                    
                    Debug.Print "szuroCellaRange.Address: "; szuroCellaRange.Address
                    
                        'hogy friss�thess�k a filterez�st, azonos�tjuk a friss�tend� oszlopot
                        Set targetAddressRng = szuroCellaRange.Offset(0, -3)
                        
                        If IsEmpty(szuroCellaRange) Then
                        
                            With Sheets(Sheet1.Name)
                            
                            'itt t�rt�nik meg a l�nyeg, a sz�r�s !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
                            .range("" & targetAddressRng & "2" & ":" & targetAddressRng & "2" & "").AutoFilter Field:=range(targetAddressRng & 1).Column

                            End With
                            
                        Else
                        
                            With Sheets(Sheet1.Name)
                            
                            'itt t�rt�nik meg a l�nyeg, a sz�r�s !!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!!
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
