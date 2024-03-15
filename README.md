# VBA-challenge
HW2

Sub ApplyToAllWorksheets()
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = "2018" Then
            homework ws
        ElseIf ws.Name = "2019" Then
            homework ws
        ElseIf ws.Name = "2020" Then
            homework ws
        End If
    Next ws
End Sub

Sub homework(ws As Worksheet)
    Dim results As Long
    Dim i As Long
    Dim lastRow As Long
    Dim ticker As String
    Dim maxTicker As String
    Dim minTicker As String
    Dim maxVolumeTicker As String
    Dim firstOpening As Double
    Dim lastClosing As Double
    Dim yearly As Double
    Dim percentchange As Double
    Dim total As Double
    Dim maxPercent As Double
    Dim minPercent As Double
    Dim maxVolume As Double
    
    lastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    results = 2
    total = 0
    
    firstOpening = Cells(2, 3).Value

    For i = 2 To lastRow
    
        total = total + Cells(i, 7).Value
    
           If Cells(i, 1).Value <> Cells(i + 1, 1).Value Then

                ticker = Cells(i, 1).Value
    
                lastClosing = Cells(i, 6).Value
    
                yearly = lastClosing - firstOpening
                
                percentchange = (yearly / firstOpening) * 100
    
                Cells(results, 9).Value = ticker
                Cells(results, 10).Value = yearly
                Cells(results, 11).Value = percentchange
                Cells(results, 12).Value = total
                results = results + 1
                    
                firstOpening = Cells(i + 1, 3).Value
                total = 0
                
            End If
            
            If Cells(i, 9).Value = Cells(i + 1, 9).Value Then
                Cells(i + 1, 9).Value = " "
            
            End If
                
            If Cells(results, 10).Value = 0 Then
              Cells(results, 10).Value = " "
              
            End If
        
    Next i

    
    For i = 2 To lastRow
    
                If Cells(i, 10) > 0 Then
                    Cells(i, 10).Interior.ColorIndex = 4
                    
                ElseIf Cells(i, 10) = 0 Then
                    Cells(i, 10).Interior.ColorIndex = xlNone
                
                Else:
                    Cells(i, 10).Interior.ColorIndex = 3
                    
                End If
    
    Next i

    maxPercent = Application.WorksheetFunction.Max(Range("K2:K" & lastRow))
    
    minPercent = Application.WorksheetFunction.Min(Range("K2:K" & lastRow))
    
    maxVolume = Application.WorksheetFunction.Max(Range("L2:L" & lastRow))
    
    
    maxTicker = Cells(Application.WorksheetFunction.Match(maxPercent, Range("K2:K" & lastRow), 0) + 1, 9).Value
    
    minTicker = Cells(Application.WorksheetFunction.Match(minPercent, Range("K2:K" & lastRow), 0) + 1, 9).Value
    
    maxVolumeTicker = Cells(Application.WorksheetFunction.Match(maxVolume, Range("L2:L" & lastRow), 0) + 1, 9).Value
    
    
    Cells(2, 16).Value = maxPercent
    Cells(2, 15).Value = maxTicker
    Cells(3, 16).Value = minPercent
    Cells(3, 15).Value = minTicker
    Cells(4, 16).Value = maxVolume
    Cells(4, 15).Value = maxVolumeTicker

    
End Sub

