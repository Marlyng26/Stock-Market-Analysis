Option Explicit

Sub analyzeData()
     Dim last_row As Long
     Dim sym As String
     Dim startPrice As Double
     Dim endPrice As Double
     Dim stockVol As Double
     Dim yearlyPriceChg As Double
     Dim yearlyPcChg As Double
     Dim largestTotVol As Double
     Dim greatestPcInc As Double
     Dim greatestPcDec As Double
     Dim i As Long
     Dim j As Integer
     Dim uFormula As String
     Dim last_tick As Long
     Dim ws As Worksheet
     
     Application.ScreenUpdating = False
     
     ' Loop through sheets
     For Each ws In Worksheets
      
         ' Get number of lines of data
         last_row = Cells(Rows.Count, 1).End(xlUp).Row
        
         ' Get first ticker symbol
         sym = Cells(2, 1).Value
         stockVol = 0
         startPrice = Cells(2, 3).Value
         If startPrice = 0 Then ' Check for value of zero to avoid divide by 0
            startPrice = 1
         End If
         
         ' Format titles
         Range("H1").Value = "Data Analysis"
         Range("I1").Value = "Ticker Symbol"
         Range("J1").Value = "Yearly Change"
         Range("K1").Value = "Percent Change"
         Range("L1").Value = "Total Stock Vol"
         Range("N1").Value = "Greatest Total Vol"
         Range("O1").Value = "Greatest % Inc"
         Range("P1").Value = "Greatest % Dec"
         Range("H1").Select
         With Selection.Font
             .Color = -4165632
         End With
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 13434879
         End With
         Selection.Font.Bold = True
          Range("I1:L1").Select
         With Selection.Font
             .Color = -4165632
         End With
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 13434879
         End With
         Selection.Font.Bold = True
         Range("N1:P1").Select
         With Selection.Font
             .Color = -4165632
         End With
         With Selection.Interior
             .Pattern = xlSolid
             .PatternColorIndex = xlAutomatic
             .Color = 13434879
         End With
         Selection.Font.Bold = True
         Columns("L:N").Select
         Selection.Style = "Comma"
         Range("O2:P2").Select
         Selection.Style = "Percent"
         
         ' Loop through data
         Range("A2").Select
         j = 2
         For i = 2 To last_row
             If Cells(i, 1).Value = sym Then                 ' Check for same ticker symbol
                 stockVol = stockVol + Cells(i, 7).Value     ' Add stock val to total
                 endPrice = Cells(i, 3).Value                ' Store what might be end price
             Else
                 yearlyPriceChg = (endPrice - startPrice)
                 yearlyPcChg = ((endPrice - startPrice) / startPrice)
                 Cells(j, 9).Value = sym                     ' Place calculated values into sheet
                 Cells(j, 10).Value = yearlyPriceChg
                 Cells(j, 10).Select
                 If ActiveCell.Value < 0 Then
                    ActiveCell.Interior.ColorIndex = 3
                 Else
                     ActiveCell.Interior.ColorIndex = 4
                 End If
                 Cells(j, 11).Value = yearlyPcChg
                 Cells(j, 11).Select
                 If ActiveCell.Value < 0 Then
                    ActiveCell.Interior.ColorIndex = 3
                 Else
                     ActiveCell.Interior.ColorIndex = 4
                 End If
                 Cells(j, 12).Value = stockVol
                 
                 If stockVol > largestTotVol Then
                     largestTotVol = stockVol
                 End If
                 If greatestPcInc < yearlyPcChg Then
                     greatestPcInc = yearlyPcChg
                 End If
                 If greatestPcDec > yearlyPcChg Then
                     greatestPcDec = yearlyPcChg
                 End If
                 sym = Cells(i, 1).Value
                 j = j + 1
                 stockVol = 0
                 startPrice = Cells(i, 3).Value
                 If startPrice = 0 Then ' Check for value of zero to avoid divide by 0
                    startPrice = 1
                End If
             End If
         Next i
         Columns("I:I").EntireColumn.AutoFit
         Columns("J:J").EntireColumn.AutoFit
         Columns("K:K").EntireColumn.AutoFit
         Columns("L:L").EntireColumn.AutoFit
         Columns("N:N").EntireColumn.AutoFit
         Columns("O:O").EntireColumn.AutoFit
         Columns("P:P").EntireColumn.AutoFit
         yearlyPriceChg = (endPrice - startPrice)
         yearlyPcChg = ((endPrice - startPrice) / startPrice)
         Cells(j, 9).Value = sym
         Cells(j, 10).Value = yearlyPriceChg
         Cells(j, 10).Select
         If ActiveCell.Value < 0 Then
            ActiveCell.Interior.ColorIndex = 3
         Else
            ActiveCell.Interior.ColorIndex = 4
         End If
         Cells(j, 11).Value = yearlyPcChg
         Cells(j, 11).Select
         If ActiveCell.Value < 0 Then
            ActiveCell.Interior.ColorIndex = 3
         Else
            ActiveCell.Interior.ColorIndex = 4
         End If
         Cells(j, 12).Value = stockVol
         Cells(2, 14).Value = largestTotVol
         Cells(2, 15).Value = greatestPcInc
         Cells(2, 16).Value = greatestPcDec
         largestTotVol = 0
         greatestPcInc = 0
         greatestPcDec = 0
         Columns("J:K").Select
         Selection.NumberFormat = "#0.00"
         Columns("L:N").Select
         Selection.NumberFormat = "_(* #,##0.0_);_(* (#,##0.0);_(* ""-""??_);_(@_)"
         Selection.NumberFormat = "_(* #,##0_);_(* (#,##0);_(* ""-""??_);_(@_)"
         Range("i4").Select
         yearlyPriceChg = 0
         yearlyPcChg = 0
         stockVol = 0
         
    If ActiveSheet.Index = Worksheets.Count Then
        Worksheets(1).Activate
    Else
        ActiveSheet.Next.Activate
        ActiveSheet.Select
    End If
    Next ws
    Application.ScreenUpdating = True
End Sub



