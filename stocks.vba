Attribute VB_Name = "Module1"
Sub changecalculations()
Dim wsdata As Worksheet
Dim lastrow As Long
Dim row As Long
Dim ticker As String
Dim openprice As Double
Dim closeprice As Double
Dim quaterlychange As Double
Dim percentagechange As Double
Dim stockvolume As Double
Dim outputrow As Long
Dim startrow As Long
Dim currentticker As String
Dim sheetnames As Variant
Dim i As Integer
Dim outputstartcolumn As Integer
Dim maxincrease As Double, maxdecrease As Double, maxvolume As Double
Dim maxdecreaseticker As String, maxvolumeticker As String
maxincrease = -100000
maxdecrease = 100000
maxvolume = 0
sheetnames = Array("Q1", "Q2", "Q3", "Q4")
For i = LBound(sheetnames) To UBound(sheetnames)
Set wsdata = ThisWorkbook.Sheets(sheetnames(i))
lastrow = wsdata.Cells(wsdata.Rows.Count, 1).End(xlUp).row
wsdata.Cells(1, 9).Value = "ticker"
wsdata.Cells(1, 10).Value = "quarterly change"
wsdata.Cells(1, 11).Value = "percent change"
wsdata.Cells(1, 12).Value = "total stock volume"
outputrow = 2
row = 2
startrow = row
Do While row <= lastrow
ticker = wsdata.Cells(row, 1).Value
openprice = wsdata.Cells(startrow, 3).Value
stockvolume = 0
If ticker = "" Or openprice = 0 Then
row = row + 1
startrow = row
GoTo continueloop
End If
Do While wsdata.Cells(row, 1).Value = ticker And row <= lastrow
closeprice = wsdata.Cells(row, 6).Value
stockvolume = stockvolume + wsdata.Cells(row, 7).Value
row = row + 1
Loop
quarterlychange = closeprice - openprice
percentagechange = (quarterlychange / openprice) * 100

wsdata.Cells(outputrow, 9).Value = ticker
wsdata.Cells(outputrow, 10).Value = Format(quarterlychange, "0.00")
wsdata.Cells(outputrow, 11).Value = Format(percentagechange, "0.00") & "%"
wsdata.Cells(outputrow, 12).Value = stockvolume
If percentagechange > maxincrease Then
maxincrease = percentagechange
maxincreaseticker = ticker
End If
If percentagechange < maxdecrease Then
maxdecrease = percentagechange
maxdecreaseticker = ticker
End If
If stockvolume > maxvolume Then
maxvolume = stockvolume
maxvolumeticker = ticker
End If
outputrow = outputrow + 1
continueloop:
startrow = row
Loop
With wsdata.Range("J2:J" & outputrow - 1)
.FormatConditions.Add Type:=xlCellValue, Operator:=xlGreater, Formula1:="=0"
.FormatConditions(1).Interior.Color = RGB(0, 255, 0)

.FormatConditions.Add Type:=xlCellValue, Operator:=xlLess, Formula1:="=0"
.FormatConditions(2).Interior.Color = RGB(255, 0, 0)
End With
wsdata.Cells(2, 15).Value = "greatest%increase"
wsdata.Cells(2, 16).Value = maxincreaseticker
wsdata.Cells(2, 17).Value = Format(maxincrease, "0.00") & "%"
wsdata.Cells(3, 15).Value = "greatest%decrease"
wsdata.Cells(3, 16).Value = maxdecreaseticker
wsdata.Cells(3, 17).Value = Format(maxdecrease, "0.00") & "%"
wsdata.Cells(4, 15).Value = "greatest total volume"
wsdata.Cells(4, 16).Value = maxvolumeticker
wsdata.Cells(4, 17).Value = maxvolume
Next i
End Sub



