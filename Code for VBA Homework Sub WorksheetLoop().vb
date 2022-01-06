Sub WorksheetLoop()

'Declare and set worksheet and workbook
Dim MainWs As Worksheet
Dim wb As Workbook

Set wb = ActiveWorkbook

'Loop through all worksheets and stocks for one year.  Not sure how to not include 2014
For Each MainWs In wb.Sheets

    'Create the column headings
        MainWs.Range("I1").Value = "Ticker"
        MainWs.Range("J1").Value = "Yearly Change"
        MainWs.Range("K1").Value = "Percent Change"
        MainWs.Range("L1").Value = "Total Stock Volume"
    
        MainWs.Range("P1").Value = "Ticker"
        MainWs.Range("Q1").Value = "Value"
        MainWs.Range("O2").Value = "Greatest % Increase"
        MainWs.Range("O3").Value = "Greatest % Decrease"
        MainWs.Range("O4").Value = "Greatest Total Volume"
    
    'Make Headers Bold
        Rows(1).Font.Bold = True
        Rows(1).VerticalAlignment = xlCenter
 Next MainWs
 
 
 'Loop through all worksheets and stocks for one year.  Not sure how to not include 2014
For Each MainWs In Worksheets
 
    'Set initial variables for calculations
    'Define Ticker variable
    Dim Ticker As String
    Ticker = " "
    Dim Ticker_volume As Double
    Ticker_volume = 0
    
    'Set new variables for prices and percent changes and calculations
    Dim beg_price As Double
    beg_price = 0
    
    Dim end_price As Double
    end_price = 0
    
    Dim yearly_price_change As Double
    yearly_price_change = 0
    
    Dim yearly_price_change_percent As Double
    yearly_price_change_percent = 0
    
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


'Set Locations for variables
Dim Summary_Table_Row As Long
Summary_Table_Row = 2

'Set row count for workbook (all sheets)
Dim Lastrow As Long

'Loop through all sheets to find last cell that is not empty
Lastrow = MainWs.Cells(Rows.Count, 1).End(xlUp).Row

'Set initial value of beginning stoick value for the first Ticker of ws
beg_price = MainWs.Cells(2, 3).Value


'Loop from the bginning of the first ws (row 2) to last row of ws 7
For i = 2 To Lastrow

    'Check if we are still on the same ticker name
    If MainWs.Cells(i + 1, 1).Value <> MainWs.Cells(i, 1).Value Then
        
        'Set ticker to starting point
        Ticker = MainWs.Cells(i, 1).Value
        
        'Calculate
        end_price = MainWs.Cells(i, 6).Value
        yearly_price_change = end_price - beg_price
        
        'set conditions for a zero value
        If beg_price <> 0 Then
            yearly_price_change_percent = (yearly_price_change / beg_price) * 100
        End If
        
        'Add to the Ticker Name Total Volume
        Total_Ticker_Volume = Total_Ticker_Volume + MainWs.Cells(i, 7).Value
        
        'Print the Ticker Name in the Summary Table, Column I
        MainWs.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print the yearly price change in the Summary Table, Column J
        MainWs.Range("J" & Summary_Table_Row).Value = yearly_price_change
        
        'Color fill yearly price change: Red for negative, and green for positive
        If (yearly_price_change > 0) Then
            MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            
        ElseIf (Yearly_Price_chnage <= 0) Then
            MainWs.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
        
        'Print the yearly price change as a percent in the Summary Table, Column K
        MainWs.Range("K" & Summary_Table_Row).Value = (CStr(yearly_price_change_percent) & "%")
        
        'Print the total stock volume in the Summary Table Column L
        MainWs.Range("L" & Summary_Table_Row).Value = Total_Ticker_Volume
        
        'Add 1 to the Summary Table Row Count
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Get next beginning price
        beg_price = MainWs.Cells(i + 1, 3).Value
        
        'Do Calculations
        If (yearly_price_change_percent > Max_Percent) Then
            Max_Percent = yearly_price_change_percent
            Max_Ticker_Name = Ticker
            
        ElseIf (yearly_price_change_percent < Min_Percent) Then
            Min_Percent = yearly_price_change_percent
            Min_Ticker_Name = Ticker
            
        End If
        
        If (Total_Ticker_Volume > Max_Volume) Then
            Max_Volume = Total_Ticker_Volume
            Max_Volume_Ticker_Name = Ticker
            
        End If
        

        
    'Else if in the next ticker name, enter new ticker stockvolume
    Else
        Total_Ticker_Volume = Total_Ticker_Volume + MainWs.Cells(i, 7).Value
    
    End If
    
Next i

        'Print values in assigned cells
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