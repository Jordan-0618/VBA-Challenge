VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Sub StockMarketInfo()

 On Error Resume Next

'Set variable for holding Ticker
 Dim Ticker As String

'Set initial variable for price change from year open to year close
 Dim Yearly_Change As Double
 Yearly_Change = 0

'Set Variable for Percent Yearly Change
 Dim Percent_Change As Double
 Percent_Change = 0

'Set initial variable for Total Stock Volume
 Dim Total_Stock_Volume As Long
 Total_Stock_Volume = 0

'Keep track of the data in the summary table
 Dim Summary_Table_Row As Integer
 Summary_Table_Row = 2

 Dim Year_Opening_Value As Double
 Year_Opening_Value = 1

 Dim Counter As Integer
 Counter = 0

'Print Ticker Header
 Range("I1").Value = "Ticker"
    
'Print Yearly Change Header
 Range("J1").Value = "Yearly Change"
    
'Print Percent Change Header
 Range("K1").Value = "Percent Change"
    
'Print Total Stock Volume
 Range("L1").Value = "Total Stock Volume"

'loop through the rows in the column
 lastrow = Cells(Rows.Count, 1).End(xlUp).Row

 For i = 2 To lastrow

'Check that we are in the same Ticker name
  If Cells(i + 1, 1).Value = Cells(i, 1).Value Then
    
    
    If Counter = 0 Then
     Year_Opening_Value = Cells(i, 3).Value
     Counter = 1
     End If
    
    'Set Ticker name
    Ticker = Cells(i, 1).Value
    
    'Add to Total Stock Volume
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value

    Else
    
    Yearly_Change = Cells(i, 6).Value - Year_Opening_Value
    
    Percent_Change = (Yearly_Change / Year_Opening_Value) * 100
    
    Total_Stock_Volume = Total_Stock_Volume + Cells(i, 7).Value
    
    'Print Ticker in Summary Table
    Range("I" & Summary_Table_Row).Value = Ticker
    
    'Print Yearly Change in Summary Table
    Range("J" & Summary_Table_Row).Value = Yearly_Change
    
    
    
    'Change Cell color for negative change to Red and positive change to Green
        If Yearly_Change > 0 Then
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
        Else
            Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
        End If
     
    'Print Yearly Percent Change in Summary Table
    Range("K" & Summary_Table_Row).Value = "%" & Percent_Change
    
    'Print Total Stock Volume in Summary Table
    Range("L" & Summary_Table_Row).Value = Total_Stock_Volume
    
    'Add Row to Summary Table
    Summary_Table_Row = Summary_Table_Row + 1
    
    'Reset Total Stock Volume to zero
     Total_Stock_Volume = 0
     
     'Reset Yearly Change Counter to Zero
    Counter = 0
 
  End If

 Next i

End Sub

