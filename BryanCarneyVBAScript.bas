Attribute VB_Name = "Module1"
Option Explicit



Sub GetDataToSummary()
    'Loop through all worksheets
    
    Dim ws As Worksheet
    
    Application.ScreenUpdating = False
    
    For Each ws In ThisWorkbook.Worksheets
    ws.Activate
           
        'Create Summary Data headers to populate data and format cells
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
        Columns("L").Select
        Selection.Style = "Percent"
        Columns("M").Select
        Selection.Style = "Comma"
        Range("O2").Select
        Selection.Style = "Percent"
        Range("O3").Select
        Selection.Style = "Percent"
        Range("O4").Select
        Selection.Style = "Comma"
        Columns("J:P").Select
        Selection.Columns.AutoFit
        Range("A1").Select
        
       
        'DeclareVariables For loop and Ticker
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
        
        Dim Ticker_Total As Double
        Ticker_Total = 0
        
        Dim SummaryRowTable As Integer
        SummaryRowTable = 1
        
        'decalre vaiable to loop through until the last row of the ticker data to capture datapoints
        Dim LastRow As Integer
        LastRow = Cells(Rows.Count, 1).End(xlUp).Row
        
       
        For i = 2 To LastRow
        
            If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                SummaryRowTable = SummaryRowTable + 1
                Ticker_Name = Cells(i, 1).Value
                Ticker_Total = Cells(i, 7).Value
                
                Range("J" & SummaryRowTable).Value = Ticker_Name
                Range("M" & SummaryRowTable).Value = Ticker_Total
                
                Range("B" & i).Value = DateSerial(Left(Cells(i, 2).Value, 4), Mid(Cells(i, 2).Value, 5, 2), Right(Cells(i, 2).Value, 2))
    
                'Convert Data into correct data type
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
                    Range("K" & SummaryRowTable).Value = QOpenBal - QCloseBal
                    Range("L" & SummaryRowTable).Value = ((QCloseBal - QOpenBal) / QOpenBal) * 100
    
                End If
                
                Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                Range("M" & SummaryRowTable).Value = Ticker_Total
        
            End If
    
        Next i
   
                'Summary table for Greatest increase,decrease and total
                
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
                Range("P3").NumberFormat = "Comma"
    Next ws

    Application.ScreenUpdating = True

End Sub

    
