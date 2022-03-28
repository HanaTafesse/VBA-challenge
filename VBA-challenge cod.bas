'Attribute VB_Name = "Module1"
Sub Stok_market()

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
'------------------------------------------------------------

Dim Ticker As String
    Ticker = " "
Dim Total_Stock_Volume As Double
    Total_Stock_Volume = 0
Dim Open_Price As Double
    Open_Price = 0
Dim Close_Price As Double
    Close_Price = 0
Dim Yearly_Change As Double
    Yearly_Change = 0
Dim Percent_Change As Double
    Percent_Change = 0
'----------------------------------------------------------
Dim MAX_Ticker As String
    MAX_Ticker = " "
Dim MIN_Ticker As String
    MIN_Ticker = " "
Dim MAX_PERCENT As Double
    MAX_PERCENT = 0
Dim MIN_PERCENT As Double
    MIN_PERCENT = 0
Dim MAX_VOLUME_TICKER As String
    MAX_VOLUME_TICKER = " "
Dim MAX_VOLUME As Double
    MAX_VOLUME = 0
'-------------------------------------------------------------
 
 Dim Summary_Table_Row As Long
 Summary_Table_Row = 2
        
Dim Lastrow As Long
Dim i As Long
Lastrow = ws.Cells(Rows.Count, 1).End(xlUp).Row
'---------------------------------------------------------

Open_Price = ws.Cells(2, 3).Value
        For i = 2 To Lastrow

If ws.Cells(i + 1, 1).Value <> ws.Cells(i, 1).Value Then
   Ticker = ws.Cells(i, 1).Value
                
  Close_Price = ws.Cells(i, 6).Value
  Yearly_Change = Close_Price - Open_Price
                
If Open_Price <> 0 Then
    Percent_Change = (Yearly_Change / Open_Price) * 100

End If
'-----------------------------------------------------------

       Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
       ws.Range("I" & Summary_Table_Row).Value = Ticker
       ws.Range("J" & Summary_Table_Row).Value = Yearly_Change
'------------------------------------------------------------

If (Yearly_Change > 0) Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
ElseIf (Yearly_Change <= 0) Then
    ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
End If
'-----------------------------------------------------------
       
       ws.Range("K" & Summary_Table_Row).Value = (CStr(Percent_Change) & "%")
       ws.Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
       Summary_Table_Row = Summary_Table_Row + 1
   Yearly_Change = 0
   Close_Price = 0
Open_Price = ws.Cells(i + 1, 3).Value
'-----------------------------------------------------------

If (Percent_Change > MAX_PERCENT) Then
                    MAX_PERCENT = Percent_Change
                    MAX_Ticker = Ticker
ElseIf (Percent_Change < MIN_PERCENT) Then
                    MIN_PERCENT = Percent_Change
                    MIN_Ticker = Ticker
End If
                       
If (Total_Stock_Volume > MAX_VOLUME) Then
    MAX_VOLUME = Total_Stock_Volume
    MAX_VOLUME_TICKER = Ticker
End If
                
    Percent_Change = 0
    Total_Stock_Volume = 0
               
Else
                
    Total_Stock_Volume = Total_Stock_Volume + ws.Cells(i, 7).Value
End If

  Next i
'---------------------------------------------------------
  
    ws.Range("Q2").Value = (CStr(MAX_PERCENT) & "%")
    ws.Range("Q3").Value = (CStr(MIN_PERCENT) & "%")
    ws.Range("P2").Value = MAX_Ticker
    ws.Range("P3").Value = MIN_Ticker
    ws.Range("Q4").Value = MAX_VOLUME
    ws.Range("P4").Value = MAX_VOLUME_TICKER
'---------------------------------------------------------
        
     Next ws
End Sub




