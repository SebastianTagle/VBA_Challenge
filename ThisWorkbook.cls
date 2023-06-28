VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub Challenge()

Dim x As Double
Dim POpen As Double
Dim PClose As Double
Dim Total_Vol As Double

For Each ws In Worksheets
    LastRow = ws.Cells(Rows.Count, "A").End(xlUp).Row
    x = 2
    Total_Vol = 0
    w = 0
    ws.Range("I1").EntireColumn.Insert
    ws.Range("I1").Value = "Ticket"
    ws.Range("J1").EntireColumn.Insert
    ws.Range("J1").Value = "Yearly Change"
    ws.Range("K1").EntireColumn.Insert
    ws.Range("K1").Value = "Percentage Change"
    ws.Range("L1").EntireColumn.Insert
    ws.Range("L1").Value = "Total Stock Volume"
        
    For i = 2 To LastRow
        
        If ws.Cells(i + 1, 1) <> ws.Cells(i, 1) Then
            p = i - w
            ws.Cells(x, 9).Value = ws.Cells(i, 1).Value
            ws.Cells(x, 12).Value = Total_Vol + ws.Cells(i, 7).Value
            POpen = ws.Cells(p, 3).Value
            PClose = ws.Cells(i, 6).Value
            ws.Cells(x, 10).Value = PClose - POpen
            ws.Cells(x, 11).Value = PClose / POpen - 1
            ws.Cells(x, 11).Style = "Percent"
                          
            x = x + 1
            w = 0
            Total_Vol = 0
        Else
            Total_Vol = Total_Vol + ws.Cells(i, 7).Value

            w = w + 1
            
        End If
                   
    Next i

    LastTicket = ws.Cells(Rows.Count, "I").End(xlUp).Row

    Max_Percent = 0
    Max_Ticket = Null
    Min_Percent = 0
    Min_Ticket = Null
    Max_Volume = 0
    Max_Volume_Ticket = Null
    
    
    For y = 2 To LastTicket
        'este if es para pintar las celdas dependiendo de si son negativos o positivos.
        If ws.Cells(y, 10).Value < 0 Then
            ws.Cells(y, 10).Interior.ColorIndex = 3
        Else
            ws.Cells(y, 10).Interior.ColorIndex = 4
        End If
        
        'este if es para ir guardando el mayor valor.
        If ws.Cells(y, 11).Value > Max_Percent Then
            Max_Percent = ws.Cells(y, 11).Value
            Max_Ticket = ws.Cells(y, 9).Value
        End If
        
        'este if es para ir guardando el menor valor.
        If ws.Cells(y, 11).Value < Min_Percent Then
            Min_Percent = ws.Cells(y, 11).Value
            Min_Ticket = ws.Cells(y, 9).Value
        End If
        
        'este if es para ir guardando el mayor volume
        If ws.Cells(y, 12).Value > Max_Volume Then
            Max_Volume = ws.Cells(y, 12).Value
            Max_Volume_Ticket = ws.Cells(y, 9).Value
        End If
        
    Next y
    
    ws.Range("P1").Value = "Ticket"
    ws.Range("Q1").Value = "Value"
    ws.Range("O2").Value = "Greatest % Income"
    ws.Range("P2").Value = Max_Ticket
    ws.Range("Q2").Value = Max_Percent
    ws.Range("Q2").Style = "Percent"
    ws.Range("O3").Value = "Greatest % Decrease"
    ws.Range("P3").Value = Min_Ticket
    ws.Range("Q3").Value = Min_Percent
    ws.Range("Q3").Style = "Percent"
    ws.Range("O4").Value = "Greatest Total Volume"
    ws.Range("P4").Value = Max_Volume_Ticket
    ws.Range("Q4").Value = Max_Volume

Next ws

End Sub



