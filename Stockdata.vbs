VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub alphabeticaltesting()

For Each ws In Worksheets

Dim ticker As String
Dim volume As Double
Dim yearlychange As Double
Dim percentchange As Double

volume = 0

ws.Range("j1").Value = "Ticker"
ws.Range("k1").Value = "Yearly change"
ws.Range("l1").Value = "Percent change"
ws.Range("m1").Value = "Total stock volume"

ws.Range("q1").Value = "Ticker"
ws.Range("r1").Value = "Value"
ws.Range("p2").Value = "Greatest percent Increase"
ws.Range("p3").Value = "Greatest percent Decrease"
ws.Range("p4").Value = "Greatest Total Volume"

Dim summarytable As Integer
Dim openvalue As Long
Dim i As Long
summarytable = 2
openvalue = 2

lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For i = 2 To lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
ticker = ws.Cells(i, 1).Value
volume = volume + ws.Cells(i, 7).Value
yearopen = ws.Range("c" & openvalue).Value
yearclose = ws.Range("f" & i).Value
yearlychange = yearclose - yearopen
percentchange = yearlychange / yearopen

ws.Range("j" & summarytable).Value = ticker
ws.Range("m" & summarytable).Value = volume
ws.Range("k" & summarytable).Value = yearlychange
ws.Range("l" & summarytable).Value = percentchange
ws.Range("l" & summarytable).NumberFormat = "0.00%"

If yearlychange > 0 Then
ws.Range("k" & summarytable).Interior.ColorIndex = 4
Else
ws.Range("k" & summarytable).Interior.ColorIndex = 3
End If

openvalue = i + 1
summarytable = summarytable + 1
volume = 0

Else

volume = volume + ws.Cells(i, 7).Value

End If

Next i

Next ws

End Sub

Function getmaxvalue() As Range
Dim i As Double
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
For i = 2 To lastrow
If Range("m" & i).Value > i Then
i = Range("m" & i).Value
End If
Next
getmaxvalue = i

End Function
