Sub TareaVBA()
 
 'Define Variables
   Dim Ticker As String
   Dim Year_Change As Double
   Dim Percent_Change As Double
   Dim Stock_Vol As Double
   Dim Open_Price As Double
   Dim Close_Price As Double
   Dim Last_Row As Long
   Dim i As Long
   Dim j As Long
   Dim Count1 As Long
   Dim Count2 As Long
   Dim WS As Worksheet
        
 'Apply to every worksheet in this workbook
  For Each WS In Worksheets
   
 'Set initial values for variables
   Stock_Vol = 0
   Count1 = 2
   Count2 = 2
        
 'Select all rows that contains value
   Last_Row = WS.Cells(Rows.count, 1).End(xlUp).Row
 
 'Create Headers
 'Table 1 Ticker Summary
   WS.Range("J1").Value = "Ticker"
   WS.Range("K1").Value = "Yearly Change"
   WS.Range("L1").Value = "% Change"
   WS.Range("M1").Value = "Total Stock Volume"
   
 'Table 2 Greatest Changes
   WS.Range("P1").Value = "Max Changes"
   WS.Range("Q1").Value = "Ticker"
   WS.Range("R1").Value = "Value"
   WS.Range("P2").Value = "Greatest % Increase"
   WS.Range("P3").Value = "Greatest % Decrease"
   WS.Range("P4").Value = "Greatest Total Volume"
   

 'First Loop to create the Ticker Summary Table
   For i = 2 To Last_Row
  
   Stock_Vol = Stock_Vol + WS.Cells(i, 7).Value

   If WS.Cells(i + 1, 1).Value <> WS.Cells(i, 1).Value Then
    Ticker = WS.Cells(i, 1).Value
    WS.Range("J" & Count1).Value = Ticker
    WS.Range("M" & Count1).Value = Stock_Vol
    WS.Range("M" & Count1).NumberFormat = "#,##0"
        
    Open_Price = WS.Range("C" & Count2).Value
    Close_Price = WS.Cells(i, 6).Value
    Year_Change = Close_Price - Open_Price
    WS.Range("K" & Count1).Value = Year_Change
    WS.Range("K" & Count1).NumberFormat = "0.00"
    
    If Year_Change >= 0 Then
    WS.Range("K" & Count1).Interior.ColorIndex = 4
    Else
    WS.Range("K" & Count1).Interior.ColorIndex = 3
    End If
    
           
   'Verify Open_Price = 0
    If Open_Price = 0 Then
    Percent_Change = 0
    WS.Range("L" & Count1).Value = Percent_Change
    WS.Range("L" & Count1).NumberFormat = "0.00%"
        
    Else
    Percent_Change = Year_Change / Open_Price
    WS.Range("L" & Count1).Value = Percent_Change
    WS.Range("L" & Count1).NumberFormat = "0.00%"
    
    
    End If
    
  'Reset variables
   Stock_Vol = 0
   Count1 = Count1 + 1
   Count2 = i + 1
        
   End If
  Next i
 
 
 'Second loop for Bonus
  Last_Row = WS.Cells(Rows.count, 12).End(xlUp).Row
 
   For j = 2 To Last_Row
   
   If WS.Range("L" & j).Value > WS.Range("R2").Value Then
   WS.Range("R2").Value = WS.Range("L" & j).Value
   WS.Range("Q2").Value = WS.Range("J" & j).Value
   WS.Range("R2").NumberFormat = "0.00%"
   End If
   
   If WS.Range("L" & j).Value < WS.Range("R3").Value Then
   WS.Range("R3").Value = WS.Range("L" & j).Value
   WS.Range("Q3").Value = WS.Range("J" & j).Value
   WS.Range("R3").NumberFormat = "0.00%"
   End If
   
   If WS.Range("M" & j).Value > WS.Range("R4").Value Then
   WS.Range("R4").Value = WS.Range("M" & j).Value
   WS.Range("Q4").Value = WS.Range("J" & j).Value
   WS.Range("R4").NumberFormat = "#,##0"
   End If
   
  Next j
 
 Next WS

End Sub
