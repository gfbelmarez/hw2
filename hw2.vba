Sub aggregations()
    Dim lastrow As Double
    Dim tickercolumn As Double
    Dim volumecolumn As Double
    Dim otickercolumn As Integer
    Dim ovolumecolumn As Integer
    Dim volumetotal As Double
    Dim ticker As String
    Dim firstprice As Double
    Dim lastprice As Double
    
    otickercolumn = 9
    ovolumecolumn = 10
    oyearlychangecolumn = 11
    opercentchangecolumn = 12
    
    Cells(1, otickercolumn).Value = "ticker"
    Cells(1, ovolumecolumn).Value = "volume"
    Cells(1, oyearlychangecolumn).Value = "yearly change"
    Cells(1, opercentchangecolumn).Value = "percent change"
    
    tickercolumn = 1
    volumecolumn = 7
    opencolumn = 3
    currow = 2
    volumetotal = 0
    closecolumn = 6
    firstprice = Cells(2, opencolumn).Value
    lastrow = Cells(Rows.Count, 2).End(xlUp).Row
    
    
    For i = 2 To lastrow
        ticker = Cells(i, tickercolumn).Value
        volumetotal = volumetotal + Cells(i, volumecolumn).Value
        If (ticker <> Cells(i + 1, tickercolumn).Value) Then
            
            lastprice = Cells(i, closecolumn)
            
            Cells(currow, otickercolumn).Value = ticker
            Cells(currow, ovolumecolumn).Value = volumetotal
            Cells(currow, oyearlychangecolumn).Value = lastprice - firstprice
            If (firstprice = 0) Then
                Cells(currow, opercentchangecolumn).Value = 0
            Else
                Cells(currow, opercentchangecolumn).Value = FormatPercent((lastprice - firstprice) / firstprice)
            End If
            If (Cells(currow, oyearlychangecolumn).Value < 0) Then
                Cells(currow, oyearlychangecolumn).Interior.Color = RGB(255, 0, 0)
            Else
                Cells(currow, oyearlychangecolumn).Interior.Color = RGB(0, 255, 0)
            End If
            
            currow = currow + 1
            volumetotal = 0
            firstprice = Cells(i + 1, opencolumn)
            
        End If
    Next i
End Sub

Sub processData()
    Dim ws As Worksheet
    For Each ws In ActiveWorkbook.Worksheets
        ws.Activate
        Call aggregations
        Call maxData
    Next ws
End Sub

Sub maxData()
    Dim lastrow As Integer
    Dim greatestVolume As Double
    Dim greatestPercent As Double
    Dim greatestChange As Double

    lastrow = Cells(Rows.Count, 9).End(xlUp).Row
    With Application.WorksheetFunction
        Cells(1, 14).Value = "Greatest Volume"
        greatestVolume = .Max(Range("J2:J" & lastrow))
        Cells(1, 15).Value = Cells(.Match(greatestVolume, Range("J1:J" & lastrow), 0), 9)
        Cells(1, 16).Value = greatestVolume
        
        Cells(2, 14).Value = "Greatest Yearly Change"
        greatestChange = .Max(Range("K2:K" & lastrow))
        Cells(2, 15).Value = Cells(.Match(greatestChange, Range("K1:K" & lastrow), 0), 9)
        Cells(2, 16).Value = greatestChange
        
        Cells(3, 14).Value = "Greatest % Change"
        greatestPercent = .Max(Range("L2:L" & lastrow))
        Cells(3, 15).Value = Cells(.Match(greatestPercent, Range("L1:L" & lastrow), 0), 9)
        Cells(3, 16).Value = FormatPercent(greatestPercent / 100)
    End With
End Sub
