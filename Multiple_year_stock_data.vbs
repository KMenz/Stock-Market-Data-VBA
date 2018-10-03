VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Sub StockMarketHard():

'Set Worksheet Variable and Being Looping ws
'Dim ws As Worksheet

For Each ws In Worksheets


'Set Variable For Holding Ticker
Dim Ticker As String

'Set Initial Variable For Ticker Total
Dim Ticker_Total As Double

'Keep Track Of Summary Table Row and Total Row
Dim Summary_Table_Row As Long
Summary_Table_Row = 2


'Set Titles For Summary Table
ws.Range("I1") = "Ticker Name"
ws.Range("J1") = "Ticker Total"
ws.Range("K1") = "Yearly Change"
ws.Range("L1") = "Percent Change"


'Create LastRow Variable for Loops
Dim lastrow As Long
lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row


'Create Open and Close Variables For Holding Prices and Price Difference and % Change and loop
Dim OpenPrice As Double
Dim ClosePrice As Double
Dim PriceDiff As Double
Dim PercentChange As Double
Dim i As Long

'Set Initial Open Price
OpenPrice = ws.Cells(2, 3).Value

'Loop Through Tickers and check to see if ticker equals the one above
    For i = 2 To lastrow
    
        
    'Set If Statement To Find If We Are In Same Ticker Symbol
        If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
              
        'Set New Ticker Name and Add To Ticker Total
        Ticker = ws.Cells(i, 1).Value
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
        
        'Print Ticker Name In Summary Table
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Print Total Amount To Summary Table
       ws.Range("J" & Summary_Table_Row).Value = Ticker_Total
       
       'Determine Close Price
       ClosePrice = ws.Cells(i, 6).Value
        
    'Determine Price Diff
    PriceDiff = ClosePrice - OpenPrice
    
    'Print Price Diff
    ws.Range("K" & Summary_Table_Row).Value = PriceDiff
    
    'Determine Percent Change
        If OpenPrice = 0 Then
        PercentChange = 0
        
        Else
        PercentChange = PriceDiff / OpenPrice
        
        End If
    
        'Print Percent Change and Change Format
        ws.Range("L" & Summary_Table_Row).Value = PercentChange
        ws.Range("L" & Summary_Table_Row).NumberFormat = "0.00%"
          
        'Add 1 To Summmary Table Row
        Summary_Table_Row = Summary_Table_Row + 1
        
        'Reset Total To 0
        Ticker_Total = 0
        
        'Reset Open Price To New Price
        OpenPrice = ws.Cells(i + 1, 3)
        
        'Otherwise If They Equal, Add to Ticker Total
        Else
        Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                
        End If
        
   Next i
   
   
'Set Conditional Formatting On Yearly Change

'Determine Last Row Of Summary Table
sumlastrow = ws.Cells(Rows.Count, 11).End(xlUp).Row

For J = 2 To sumlastrow

    If ws.Cells(J, 11).Value > 0 Then
    ws.Cells(J, 11).Interior.ColorIndex = 10
    
    ElseIf ws.Cells(J, 11).Value < 0 Then
    ws.Cells(J, 11).Interior.ColorIndex = 3
    
    End If
               
Next J


'Set Titles For Greatest Increase, Decrease, and Volume
ws.Range("O1").Value = "Ticker Name"
ws.Range("P1").Value = "Amount"

ws.Range("N2") = "Greatest % Increase"
ws.Range("N3") = "Greatest % Decrease"
ws.Range("N4") = "Greatest Volume"

'Look Through Each Row and Find The Greatest Values
For Z = 2 To sumlastrow

If ws.Cells(Z, 12).Value = Application.WorksheetFunction.Max(ws.Range("L2:L" & sumlastrow)) Then
    ws.Range("O2").Value = ws.Cells(Z, 9).Value
    ws.Range("P2").Value = ws.Cells(Z, 12).Value
    ws.Range("P2").NumberFormat = "0.00%"
    
ElseIf ws.Cells(Z, 12).Value = Application.WorksheetFunction.Min(Range("L2:L" & sumlastrow)) Then
    ws.Range("O3").Value = ws.Cells(Z, 9).Value
    ws.Range("P3").Value = ws.Cells(Z, 12).Value
    ws.Range("P3").NumberFormat = "0.00%"
    
ElseIf ws.Cells(Z, 10).Value = Application.WorksheetFunction.Max(Range("J2:J" & sumlastrow)) Then
    ws.Range("O4").Value = ws.Cells(Z, 9).Value
    ws.Range("P4").Value = ws.Cells(Z, 10).Value
    
    End If
    
    
Next Z
    
    
Next ws
                      
End Sub


