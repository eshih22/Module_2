Sub stockanalysis():
    For Each ws In Worksheets
    
       Dim tsv As Double
       Dim ticker As String
       Dim summaryrow As Integer
       Dim openprice As Double
                 
       tsv = 0
       summaryrow = 2
       
       ws.Range("I1").Value = "Ticker"
       ws.Range("J1").Value = "Yearly Change"
       ws.Range("K1").Value = "Percent Change"
       ws.Range("L1").Value = "Total Stock Volume"
           
       lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
       ws.Range("K:K").NumberFormat = "0.00%"
       ws.Range("L:L").NumberFormat = "#,##0"
       ws.Range("Q2:Q3").NumberFormat = "0.00%"
       ws.Range("Q4").NumberFormat = "#,##0"
       
       For i = 2 To lastrow
           If ws.Cells(i, 1).Value <> ws.Cells(i + 1, 1).Value Then
               ticker = ws.Cells(i, 1).Value
               tsv = tsv + ws.Cells(i, 7).Value
               closeprice = ws.Cells(i, 6).Value
                   
                  
               ws.Range("I" & summaryrow).Value = ticker
               ws.Range("L" & summaryrow).Value = tsv
               ws.Range("J" & summaryrow).Value = (openprice - closeprice)
               ws.Range("K" & summaryrow).Value = (Range("J" & summaryrow).Value / openprice)
                   
               tsv = 0
               summaryrow = summaryrow + 1
                   
           ElseIf ws.Cells(i, 1).Value <> ws.Cells(i - 1, 1).Value Then
               openprice = ws.Cells(i, 3).Value
               
           Else
               tsv = tsv + ws.Cells(i, 7).Value
               
           End If
                 
       Next i
       
       Dim maxchange As Double
       Dim maxticker As String
       Dim minchange As Double
       Dim minticker As String
       Dim maxvolticker As String
       Dim maxvol As Double
       
       maxchange = ws.Range("K2").Value
       minchange = ws.Range("K2").Value
       maxvol = ws.Range("L2").Value
                         
       ws.Range("P1").Value = "Ticker"
       ws.Range("Q1").Value = "Value"
       ws.Range("O2").Value = "Greatest % Increase"
       ws.Range("O3").Value = "Greatest % Decrease"
       ws.Range("O4").Value = "Greatest Total Volume"
       
       For i = 2 To lastrow
           If ws.Cells(i, 11).Value > maxchange Then
           maxticker = ws.Cells(i, 9).Value
           maxchange = ws.Cells(i, 11).Value
                             
           ElseIf ws.Cells(i, 11).Value < minchange Then
           minticker = ws.Cells(i, 9).Value
           minchange = ws.Cells(i, 11).Value
           
           ElseIf ws.Cells(i, 12).Value > maxvol Then
           maxvolticker = ws.Cells(i, 9).Value
           maxvol = ws.Cells(i, 12).Value
           
           End If
           
           If ws.Cells(i, 11) > 0 Then
           ws.Cells(i, 11).Interior.ColorIndex = 4
           
           ElseIf ws.Cells(i, 11) < 0 Then
          ws.Cells(i, 11).Interior.ColorIndex = 3
           
           End If
       
       Next i
    ws.Range("P2").Value = maxticker
    ws.Range("Q2").Value = maxchange
    ws.Range("P3").Value = minticker
    ws.Range("Q3").Value = minchange
    ws.Range("P4").Value = maxvolticker
    ws.Range("Q4").Value = maxvol
    ws.Columns("A:Q").AutoFit
    
    Next ws
      
End Sub

