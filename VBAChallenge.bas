Attribute VB_Name = "Module1"
Sub Yearly_Stock_Analysis()

Dim InfoSht As Worksheet
    
For Each InfoSht In ActiveWorkbook.Worksheets
    InfoSht.Activate

    'Set Up Constants
    Const TICKCOL As Integer = 1
    Const VOLCOL As Integer = 7
    Const OPNCOL As Integer = 3
    Const CLSCOL As Integer = 6
    
    'Set Up Variables
    Dim OpenPrice, ClosePrice, YrlyChng, YrlyPrct As Double
    Dim TickerRow, InputRow, OutputRow As Integer
    Dim Volume, LastRow As Long
    Dim TickName As String
               
    LastRow = Cells(Rows.Count, 1).End(xlUp).Row
    
    'Start Counters
    Volume = 0
    OutputRow = 2
    InputRow = 2
    
    'Set header
    InfoSht.Cells(1, 9).Value = "Ticker"
    InfoSht.Cells(1, 10).Value = "Yearly Change"
    InfoSht.Cells(1, 11).Value = "Percent Change"
    InfoSht.Cells(1, 12).Value = "Total Stock Volume"
    
        For InputRow = 2 To LastRow
          TickName = Cells(InputRow, TICKCOL).Value
          Volume = Volume + Cells(InputRow, VOLCOL).Value
          
            'If First row of current ticker
           If Cells(InputRow - 1, TICKCOL).Value <> TickName Then
                    OpenPrice = Cells(InputRow, OPNCOL).Value
          End If
          
            'If last row of current ticker
            If Cells(InputRow + 1, TICKCOL).Value <> TickName Then
                'Calculations
                ClosePrice = Cells(InputRow, CLSCOL).Value
                YrlyChng = ClosePrice - OpenPrice
                If (OpenPrice = 0) Then
                    YrlyPrct = 0
                Else
                    YrlyPrct = ((ClosePrice / OpenPrice) - 1)
               End If
                
                'Writing
                Cells(OutputRow, 9).Value = TickName
                Cells(OutputRow, 10).Value = YrlyChng
                Cells(OutputRow, 11).Value = YrlyPrct
                Cells(OutputRow, 11).NumberFormat = "0.00%"
                Cells(OutputRow, 12).Value = Volume
                
                'Counters Reset
                OutputRow = OutputRow + 1
                Volume = 0

           End If
           Next InputRow
 Next InfoSht
 End Sub


