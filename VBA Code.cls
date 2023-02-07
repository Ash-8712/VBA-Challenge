VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StocksLoop()

    'set worksheet as object variable
    Dim headers() As Variant
    Dim MainWs As Worksheet
    Dim wb As Workbook
    
    Set wb = ActiveWorkbook
    
    'fix header info
    headers() = Array("Ticker", "Date", "Open", "High", "Low", "Close", "Volume", " ", "Ticker", "Yearly-Change", "Percent_Change", "Stock_Volume", " ", " ", " ", "Ticker", "Value")
    
    For Each MainWs In wb.Sheets
        With MainWs
        .Rows(1) = " "
        For i = LBound(headers()) To UBound(headers())
        .Cells(1, 1 + i).Value = headers(i)
        
        Next i
        .Rows(1).Font.Bold = True
        .Rows(1).VerticalAlignment = xlCenter
        .Rows(1).Columns.AutoFit
        End With
    Next MainWs
    
    '--------------------------------------------------------
    
    'loop through all the worksheet in the workbook
    For Each MainWs In Worksheets
    
    'delclare the strings and doubles
        Dim Ticker_Name As String
        Ticker_Name = " "
        Dim Total_Ticker_Volume As Double
        Total_Ticker_Volume = 0
        Dim Beg_Price As Double
        Beg_Price = 0
        Dim End_Price As Double
        End_Price = 0
        Dim Yearly_Price_Change As Double
        Yearly_Price_Change = 0
        Dim Yearly_Price_Change_Percent As Double
        Yearly_Price_Change_Percent = 0
        Dim Max_Ticker_Name As String
        Max_Ticker_Name = " "
        Dim Min_Ticker_Name As String
        Min_Ticker_Name = " "
        Dim Max_Percent As Double
        Max_Percent = 0
        Dim Min_Percent As Double
        Min_Percent = 0
        Dim Max_Volume_Ticker_Name As String
        Max_Volume_Ticker_Name = " "
        Dim Max_Volume As Double
        Max_Volume = 0
        
    '----------------------------------------------------------
    
    'location for my declarations
    Dim summary_Table_Row As Long
    summary_Table_Row = 2
    
    'find lastrow
    Dim LastRow As Long
    LastRow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row
    'finished looping through wb for last rows
    
    'find the begining stock value for the first ticker
    Beg_Price = MainWs.Cells(2, 3).Value
    
    'now I need to create a loop to scan all worksheets from row 2 until the the 6th worksheet
    For i = 2 To LastRow
    
        'check the ticker names for differences
        If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then
        
        'set ticker names
        Ticker_Name = MainWs.Cells(i, 1).Value
        
        'Calculate
        End_Price = MainWs.Cells(i, 6).Value
        Yearly_Price_Change = End_Price - Beginning_Price
        
        'I need a zero value
        If Beg_Price <> 0 Then
            Yearly_Price_Change_Percent = (Yearly_Price_Change / Beg_Price) * 100
        End If
        
        'add up the ticker total volume
        Total_Ticker_Volume = Total_Ticker_Volume + MainWs.Cells(i, 7).Value
        
        'Print the ticker name in the summary table column i
        MainWs.Range("I" & summary_Table_Row).Value = Ticker_Name
        
        'print the yearly price change in the summary table , column k
        MainWs.Range("J" & summary_Table_Row).Value = Yearly_Price_Change
        
        'color fill yearly price change: red for negative, green for positive
        If (Yearly_Price_Change > 0) Then
            MainWs.Range("J" & summary_Table_Row).Interior.ColorIndex = 4
            
        ElseIf (Yearly_Price_Change < 0) Then
            MainWs.Range("J" & summary_Table_Row).Interior.ColorIndex = 3
            
        End If
        
        'put the yearly price change  as a percent in column k
        MainWs.Range("K" & summary_Table_Row).Value = (CStr(Yearly_Price_Change_Percent) & "%")
        
        'put total stock volume in summary table
        MainWs.Range("L" & summary_Table_Row).Value = Total_Ticker_Volume
        
        'add one to the summary row count
        summary_Table_Row = summary_Table_Row + 1
        
        'get next beginning price
        Beg_Price = MainWs.Cells(i + 1, 3).Value
        
        'if statement to figure out max and min calculations
        If (Yearly_Price_Change_Percent > Max_Percent) Then
            Max_Percent = Yearly_Price_Change_Percent
            Max_Ticer_Name = Ticker_Name
            
        ElseIf (Yearly_Price_Change_Percent < Min_Percent) Then
            Min_Percent = Yearly_Price_Change_Percent
            Min_Ticker_Name = Ticker_Name
        End If
        
        If (Total_Ticker_Volume > Max_Volume) Then
            Max_Volume = Total_Ticker_Volume
            Max_Volume_Ticker_Name = Ticker_Name
        End If
        
        'Reset the values
        Yearly_Price_Change_Percent = 0
        Total_Ticker_Volume = 0
        
    Else
    
        Total_Ticker_Volume = Total_Ticker_Volume + MainWs.Cells(i, 7).Value
        
    End If
    
    Next i
    
    'put the values in the cells
    MainWs.Range("Q2").Value = (CStr(Max_Percent) & "%")
    MainWs.Range("Q3").Value = (CStr(Min_Percent) & "%")
    MainWs.Range("P2").Value = Max_Ticker_Name
    MainWs.Range("P3").Value = Min_Ticker_Name
    MainWs.Range("Q4").Value = Max_Volume
    MainWs.Range("O2").Value = "Greatest % Increase"
    MainWs.Range("O3").Value = "Greatest % Decrease"
    MainWs.Range("O4").Value = "Greatest Total Volume"
    
    Next MainWs
    
    
End Sub
