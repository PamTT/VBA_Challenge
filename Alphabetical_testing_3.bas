Attribute VB_Name = "Module1"
Sub Alphabetical_Testing()

    Dim Ticker_Name As String
    Dim Yearly_Change As Double
    Dim Percent_Change As Double
    Dim Volume_Total As LongLong
    Dim Summary_Table_Row As Integer
    Dim Year_Begin_Price As Double
    Dim Year_Close_Price As Double
    
      
 
'Print Headers
        Range("J1").Value = "Ticker"
        Range("K1").Value = "Yearly_Change"
        Range("L1").Value = "Percent_Change"
        Range("M1").Value = "Volume_Total"

'Setup Starting and End Location and number format
    Year_Begin_Price = Cells(2, 3).Value
    lastrow = Cells(Rows.Count, 1).End(xlDown).Row
    Columns(12).NumberFormat = "##,##0.00%"
    Summary_Table_Row = 2
    
 'For Loop
    For i = 2 To lastrow
        If Cells(i - 1, 1).Value <> Cells(i, 1).Value And i <> 2 Then
            Year_End_Price = Cells(i - 1, 6).Value
            Ticker_Name = Cells(i - 1, 1).Value
            
            Range("J" & Summary_Table_Row).Value = Ticker_Name
            Range("M" & Summary_Table_Row).Value = Volume_Total
            Yearly_Change = Year_End_Price - Year_Begin_Price
            
            If Year_Begin_Price <> 0 Then
                Percent_Change = (Year_End_Price / Year_Begin_Price - 1)
            Else
                Percent_Change = -10000
            End If
            
            Range("K" & Summary_Table_Row).Value = Yearly_Change
            Range("L" & Summary_Table_Row).Value = Percent_Change
            
            If Yearly_Change > 0 Then
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                Range("K" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
                
            Summary_Table_Row = Summary_Table_Row + 1
            Year_Begin_Price = Cells(i, 3).Value
            Volume_Total = 0
        End If
        Volume_Total = Volume_Total + Cells(i, 7).Value
    Next i
End Sub
    

 

