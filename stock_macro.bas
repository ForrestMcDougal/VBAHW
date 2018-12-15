Attribute VB_Name = "Module1"
Sub stock_macro():
Attribute stock_macro.VB_ProcData.VB_Invoke_Func = " \n14"

    'Find the number of worksheets in workbook
    Dim WS_COUNT As Integer
    Dim k As Integer
    WS_COUNT = ActiveWorkbook.Worksheets.Count

    'Loop through all of the worksheets in workbook
    'Note: this is not the most efficient way to do this, since it is 
    'purposefully modular.
    For k = 1 To WS_COUNT
        Sheets(k).Select
        Call easy_macro
        Call moderate_macro
        Call hard_macro
    Next k
End Sub

Sub easy_macro():
    Dim ticker As String
    Dim NUM_COLUMS_A, i As Long
    Dim placement As Integer
    NUM_COLUMNS_A = Range("A" & Rows.Count).End(xlUp).Row
    
    ticker = Cells(2, "A").Value
    placement = 2
    Cells(1, "I").Value = "Ticker"
    Cells(1, "J").Value = "Total Stock Volume"
    Cells(2, "I") = ticker
    Columns("J").NumberFormat = "#,##0"
    
    'Loop through all rows, adding volume to ticker if same ticker, new row and new ticker if not
    Cells(2, "J").Value = 0
    For i = 2 To NUM_COLUMNS_A
        If (Cells(i, "A").Value = ticker) Then
            Cells(placement, "J").Value = Cells(placement, "J").Value + Cells(i, "G").Value
        Else
            placement = placement + 1
            ticker = Cells(i, "A").Value
            Cells(placement, "I").Value = ticker
            Cells(placement, "J").Value = Cells(i, "G").Value
        End If
    Next i
End Sub

Sub moderate_macro():
    Dim NUM_COLUMS_A, NUM_COLUMNS_P, i, placement As Long
    Dim stock_open As Double
    
    'Find number of columns in A and I, A for differences and I for coloring
    NUM_COLUMNS_A = Range("A" & Rows.Count).End(xlUp).Row
    NUM_COLUMNS_I = Range("I" & Rows.Count).End(xlUp).Row
    
    ticker = Cells(2, "A").Value
    placement = 2
    stock_open = Cells(2, "C").Value
    
    Cells(1, "K").Value = "Yearly Change"
    Cells(1, "L").Value = "Percent Change"
    
    For i = 2 To NUM_COLUMNS_A
        If (Cells(i, "A").Value <> ticker) Then
            ticker = Cells(i, "A").Value
            stock_close = Cells(i - 1, "F").Value
            Cells(placement, "K").Value = stock_close - stock_open

            'Handle divide by 0 error if volume is 0 for the year
            If (stock_open > 0) Then
                Cells(placement, "L").Value = Cells(placement, "K") / stock_open
            Else
                Cells(placement, "L").Value = 0
            End If

            placement = placement + 1
            stock_open = Cells(i, "C").Value
        End If
    Next i
    Cells(placement, "K").Value = stock_open - Cells(i, "F").Value
    Columns("K").NumberFormat = "General"
    
    'Color cells on difference
    For i = 2 To NUM_COLUMNS_I
        If (Cells(i, "L").Value >= 0) Then
            Cells(i, "L").Interior.Color = vbGreen
        Else
            Cells(i, "L").Interior.Color = vbRed
        End If
        If (Cells(i, "K").Value >= 0) Then
            Cells(i, "K").Interior.Color = vbGreen
        Else
            Cells(i, "K").Interior.Color = vbRed
        End If
    Next i
    
    Columns("L").NumberFormat = "0.00%"
    
End Sub

Sub hard_macro():
    Cells(2, "N").Value = "Greatest % Increase"
    Cells(3, "N").Value = "Greatest % Decrease"
    Cells(4, "N").Value = "Greatest Total Volume"
    Cells(1, "O").Value = "Ticker"
    Cells(1, "P").Value = "Value"
    Cells(2, "P").Value = 0
    Cells(3, "P").Value = 0
    Cells(4, "P").Value = 0
    
    Dim NUM_COLUMNS_P As Long
    
    NUM_COLUMNS_I = Range("I" & Rows.Count).End(xlUp).Row
    
    'Loop through all tickers, checking if higher or lower for respective values
    For i = 2 To NUM_COLUMNS_I
        If (Cells(i, "L").Value > Cells(2, "P").Value) Then
            Cells(2, "P").Value = Cells(i, "L").Value
            Cells(2, "O").Value = Cells(i, "I").Value
        ElseIf (Cells(i, "L").Value < Cells(3, "P").Value) Then
            Cells(3, "P").Value = Cells(i, "L").Value
            Cells(3, "O").Value = Cells(i, "I").Value
        End If
        If (Cells(i, "J").Value > Cells(4, "P").Value) Then
            Cells(4, "P").Value = Cells(i, "J").Value
            Cells(4, "O").Value = Cells(i, "I").Value
        End If
    Next i
    
    Cells(2, "P").NumberFormat = "0.00%"
    Cells(3, "P").NumberFormat = "0.00%"
    Cells(4, "P").NumberFormat = "#,##0"
    Columns("A:P").AutoFit
    
End Sub
