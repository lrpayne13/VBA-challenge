Sub HW()


Dim ws As Worksheet
    For Each ws In Worksheets
         
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Stock Volume"
ws.Range("O2").Value = "Greatest % Increase"
ws.Range("O3").Value = "Greatest % Decrease"
ws.Range("O4").Value = "Greatest Total Volume"
ws.Range("P1").Value = "Ticker"
ws.Range("Q1").Value = "Value"

ws.Range("A2:G753001").Sort Key1:=ws.Range("A2"), Order1:=xlAscending, Header:=xlNo

Dim Ticker As String
Dim Volume As Double
Volume = 0
SumVolume = 2
Dim OpenValue As Double
Dim Difference As Double
Dim PercentChange As Double
Dim EndTable As Long

OpenValue = ws.Cells(2, 3).Value
EndTable = ws.Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To EndTable

    If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
        Ticker = ws.Cells(i, 1).Value
        Volume = Volume + ws.Cells(i, 7).Value
        CloseValue = ws.Cells(i, 6).Value
                
        Difference = CloseValue - OpenValue
        PercentChange = (Difference / OpenValue)
                                
        ws.Range("I" & SumVolume).Value = Ticker
        ws.Range("J" & SumVolume).Value = Difference
        ws.Range("K" & SumVolume).Value = PercentChange
        ws.Range("K" & SumVolume).NumberFormat = "0.00%"
        ws.Range("L" & SumVolume).Value = Volume
            If ws.Range("J" & SumVolume).Value >= 0 Then
                ws.Range("J" & SumVolume).Interior.ColorIndex = 4
            ElseIf ws.Range("J" & SumVolume).Value < 0 Then
                ws.Range("J" & SumVolume).Interior.ColorIndex = 3
            End If
                    
        SumVolume = SumVolume + 1
        
        Volume = 0
        OpenValue = ws.Cells(i + 1, 3).Value
               
    Else
        Volume = Volume + ws.Cells(i, 7).Value
    End If
        
Next i


ws.Range("Q2").Value = WorksheetFunction.Max(ws.Range("K2:K" & EndTable))
ws.Range("Q2").NumberFormat = "0.00%"
ws.Range("Q3").Value = WorksheetFunction.Min(ws.Range("K2:K" & EndTable))
ws.Range("Q3").NumberFormat = "0.00%"
ws.Range("Q4").Value = WorksheetFunction.Max(ws.Range("L2:L" & EndTable))

MaxIncreaseIndex = WorksheetFunction.Match(ws.Range("Q2").Value, ws.Range("K2:K" & EndTable), 0)
MaxDecreaseIndex = WorksheetFunction.Match(ws.Range("Q3").Value, ws.Range("K2:K" & EndTable), 0)
MaxVolumeIndex = WorksheetFunction.Match(ws.Range("Q4").Value, ws.Range("L2:L" & EndTable), 0)

ws.Range("P2").Value = ws.Cells(MaxIncreaseIndex + 1, 9).Value
ws.Range("P3").Value = ws.Cells(MaxDecreaseIndex + 1, 9).Value
ws.Range("P4").Value = ws.Cells(MaxVolumeIndex + 1, 9).Value

Next ws

End Sub
