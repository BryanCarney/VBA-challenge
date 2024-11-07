Attribute VB_Name = "Module1"
Option Explicit


Sub GetDataToSummary()

    'Trick shared by a co-worker to minimize run time on Macro
    Application.ScreenUpdating = False
    
    'Loop through all worksheets
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
           
           
        'Create Summary Data headers to populate data and format cells in each sheet to create the formatting and standarization across all sheets.
        Range("A1").Select
        ActiveCell.Range("J1").Value = "Ticker"
        ActiveCell.Range("K1").Value = "Quarterly Change"
        ActiveCell.Range("L1").Value = "Percent Change"
        ActiveCell.Range("M1").Value = "Total Stock Volume"
        ActiveCell.Range("N2").Value = "Greatest % Increase"
        ActiveCell.Range("N3").Value = "Greatest % Decrease"
        ActiveCell.Range("N4").Value = "Greatest Total Volume"
        ActiveCell.Range("O1").Value = "Ticker"
        ActiveCell.Range("P1").Value = "Value"
        With Selection
            .HorizontalAlignment = xlCenter
            .VerticalAlignment = xlBottom
        End With
        Columns("L").NumberFormat = ("0.00%")
        Columns("M").NumberFormat = ("#,000")
        Range("O2").Select
        Selection.NumberFormat = ("0.00%")
        Range("O3").Select
        Selection.Style = "Percent"
        Range("P4").Select
        Selection.NumberFormat = ("#,000")
        Columns("A:P").Select
        Selection.Columns.AutoFit
        Range("A1").Select
                 
        
         'DeclareVariables in For loop and Ticker
         Dim i As Integer
         Dim Ticker_Name As String
         
         'declare variables to grab the Min and Max Date for each ticker
         Dim MinDate As Date
         Dim MaxDate As Date
         
         'Declare variable for Opening and closing balance
         Dim QOpenBal As Double
         Dim QCloseBal As Double
             
         'Declare variable for changes
         Dim Ticker_Change As Double
         Dim Ticker_Percent As Double
         
         'Declare vaiable for Ticker Total and set it to zero value
         Dim Ticker_Total As Double
         Ticker_Total = 0
         
         'Declare vaiable for SummaryRowTable and set it to one
         Dim SummaryRowTable As Integer
         SummaryRowTable = 1
         
         'Decalre vaiable to loop through until the last row of the ticker data to capture datapoints
         Dim LastRow As Integer
         LastRow = Cells(Rows.Count, 1).End(xlUp).Row
         
         'Delcare variables for colour conditions for Quarterly Change
         Dim LastRowK As Long
         Dim cell As Range
         
         Dim LastRowL As Long
         Dim cellL As Range
         
         'For Loop to first summarize the Ticker Total Volume
         For i = 2 To LastRow

             If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                 SummaryRowTable = SummaryRowTable + 1
                 Ticker_Name = Cells(i, 1).Value
                 Ticker_Total = Cells(i, 7).Value
                 
                 Range("J" & SummaryRowTable).Value = Ticker_Name
                 Range("M" & SummaryRowTable).Value = Ticker_Total
                 
                
                 'Refortting of the data column so it could be calculated as a formal date - Not sure if completely neccessary but wanted to be sure when transferring the code to the working file to ensure the code effected the same way
                 Range("B" & i).Value = DateSerial(Left(Cells(i, 2).Value, 4), Mid(Cells(i, 2).Value, 5, 2), Right(Cells(i, 2).Value, 2))
     
                 'continuation of for loop to retrieve the first date and last Date in the series.  Written this way in case the dates were not in order.
                 MinDate = Cells(i, 2).Value
                 QOpenBal = Cells(i, 3).Value
     
     
                 MaxDate = Cells(i, 2).Value
                 QCloseBal = Cells(i, 6).Value
                             
                 
             Else
                 'Conditions to only pick up the correct values if it meet the Boolean condition True or False
                 Dim UpdatedMinMax As Boolean: UpdatedMinMax = False
                 
                 If Len(CStr(Cells(i, 2))) = 8 Then
                     Range("B" & i).Value = DateSerial(Left(Cells(i, 2).Value, 4), Mid(Cells(i, 2).Value, 5, 2), Right(Cells(i, 2).Value, 2))
                 End If
                 
                 If (Cells(i, 2)) < MinDate Then
                     MinDate = (Cells(i, 2).Value)
                     QOpenBal = Cells(i, 3).Value
                     UpdatedMinMax = True
     
                 End If
     
                 If (Cells(i, 2)) > MaxDate Then
                     MaxDate = Cells(i, 2).Value
                     QCloseBal = Cells(i, 6).Value
                     UpdatedMinMax = True
     
                 End If
     
                 If UpdatedMinMax = True Then
                     Range("K" & SummaryRowTable).Value = QCloseBal - QOpenBal
                     Range("L" & SummaryRowTable).Value = ((QCloseBal - QOpenBal) / QOpenBal)
     
                 End If
                 
                     
                 'Use conditional formatting to change the colour of the cells for Quarterly change if they reflect a positive or negative value after it generated into the cells in the code above
         
                 LastRowK = ws.Cells(ws.Rows.Count, "K").End(xlUp).Row
                 
                 For Each cell In ws.Range("K2:K" & LastRowK)
                     If IsEmpty(cell.Value) Then
                         cell.Interior.ColorIndex = xlNone
                     ElseIf cell.Value < 0 Then
                         cell.Interior.Color = RGB(255, 0, 0)
                     Else
                         cell.Interior.Color = RGB(0, 255, 0)
                     End If
                 Next cell
                 
                 For Each cellL In ws.Range("L2:L" & LastRowL)
                     If IsEmpty(cellL.Value) Then
                         cellL.Interior.ColorIndex = xlNone
                     ElseIf cellL.Value < 0 Then
                         cellL.Interior.Color = RGB(255, 0, 0)
                     Else
                         cellL.Interior.Color = RGB(0, 255, 0)
                     End If
                 
                 Next cell
                 
                 
                 Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                 Range("M" & SummaryRowTable).Value = Ticker_Total
             End If
     
         Next i
    
         'Summary table for Greatest increase, decrease and largest total
                 
         Dim j As Integer
         Dim MaxValue As Double
         Dim Max_Ticker As String
         Dim MinValue As Double
         Dim Min_Ticker As String
         Dim MaxVolume As Double
         Dim MaxVolume_Ticker As String
         
         Dim LastRow_Summary As Integer
         LastRow_Summary = Cells(Rows.Count, 12).End(xlUp).Row
     
         ' Find greatest % increase and respective ticker and place these in the summary table
     
         MaxValue = Range("L2").Value
             For j = 3 To LastRow_Summary
                 If Cells(j, 12).Value > MaxValue Then
                     MaxValue = Cells(j, 12).Value
                     Max_Ticker = Cells(j, 10).Value
                 End If
             Next j
         
         'Place values into cells
       
         Range("O2").Value = Max_Ticker
         Range("P2").Value = MaxValue
         Range("P2").NumberFormat = "0.00%"
         
         ' Find greatest % decrease and respective ticker and place these in the summary table
         MinValue = Range("L2").Value
             For j = 3 To LastRow_Summary
                 If Cells(j, 12).Value < MinValue Then
                     MinValue = Cells(j, 12).Value
                     Min_Ticker = Cells(j, 10).Value
                 End If
             Next j
         
         'Place values into cells
         Range("O3").Value = Min_Ticker
         Range("P3").Value = MinValue
         Range("P3").NumberFormat = "0.00%"
         
         ' Find greatest total increase and respective ticker and place these in the summary table
         MaxVolume = Range("M2").Value
             For j = 3 To LastRow_Summary
                 If Cells(j, 13).Value > MaxVolume Then
                     MaxVolume = Cells(j, 13).Value
                     MaxVolume_Ticker = Cells(j, 10).Value
                 End If
             Next j
         
         'Place values into cells
       
         Range("O4").Value = MaxVolume_Ticker
         Range("P4").Value = MaxVolume

    Next ws

    Application.ScreenUpdating = True

End Sub

    
