Sub VBAStocks():

 For Each ws In Worksheets

 Dim i As Long
 Dim j As Long
 Dim start As Long
 Dim lastrow As Long
 Dim ticker As String
 Dim total As Double
 Dim yearlychange As Double
 Dim percentchange As Double
 Dim openprice As Double
 Dim closeprice As Double
 
 ws.Range("I1").Value = "Ticker"
 ws.Range("J1").Value = "Yearly Change"
 ws.Range("K1").Value = "Percent Change"
 ws.Range("L1").Value = "Total Stock Value"

 total = 0
 j = 2
 start = 2

 lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row

 For i = 2 To lastrow
         
    If ws.Range("A" & i + 1).Value = ws.Range("A" & i).Value Then
        total = total + ws.Range("G" & i).Value
    Else
    ticker = ws.Range("A" & i).Value
    openprice = ws.Range("C" & start)
    closeprice = ws.Range("F" & i)
    yearlychange = closeprice - openprice

         If openprice = 0 Then
            percentchange = 0
         Else
            percentchange = yearlychange / openprice
         End If
     
         ws.Range("I" & j).Value = ticker
         ws.Range("L" & j).Value = total + ws.Range("G" & i).Value
         ws.Range("J" & j).Value = yearlychange
         ws.Range("K" & j).Value = percentchange
         ws.Range("K" & j).NumberFormat = "0.00%"
         
         If ws.Range("J" & j).Value > 0 Then
            ws.Range("J" & j).Interior.ColorIndex = 4
         Else
            ws.Range("J" & j).Interior.ColorIndex = 3
         End If

         j = j + 1
         total = 0
         start = i + 1
         
     End If
     
 Next i
 
 Next ws
 
End Sub
