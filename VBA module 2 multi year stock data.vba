Sub Assignment_2_VBA()

'Looping for the tickers and aggrehating total stock volume

Dim wb As Workbook
Dim sht As Worksheet
Dim rg As Range

Set wb = ThisWorkbook
Set sht = wb.Worksheets("2018")
Set rg = sht.Range("I1")

sht.Range("A:A").AdvancedFilter Action:=xlFilterCopy, CopytoRange:=rg, Unique:=True

Dim EndRow As Long

EndRow = rg.End(xlDown).Row

For x = 2 To EndRow

sht.Cells(x, 12) = WorksheetFunction.SumIf(sht.Range("A:A"), sht.Cells(x, 9), sht.Range("G:G"))
 Next x
 

End Sub

__________________________________________________________________________________________________________________

Sub ChangeInPrice()
 
' Calculating change in opening and closing price
Dim wb As Workbook
Dim sht As Worksheet
Dim rg As Range

Set wb = ThisWorkbook
Set sht = wb.Worksheets("2018")
Set rg = sht.Range("A1")

Dim EndRow As Long

EndRow = rg.End(xlDown).Row
 
For i = 2 To EndRow
 
 If sht.Cells(i + 1, i).Value <> i Then
  sht.Cells(x, 10).Value = Range("G") - Range("C")

 
 End If
 
 
 Next i

End Sub