# VBA-challenge

'################################################################################################################################################
In the below part of the code, I was removing the screen updating prior to the code to minimize of the MACRO's run time.  This was a helpful tool provided by a co-worker
I mentioned that the testing data was taking a long time to complete while working on the code and they suggested this and it really assisted in minimizing the turnaround time. 

    Code Extract:
    'Trick shared by a co-worker to minimize run time on Macro
    Application.ScreenUpdating = False

'################################################################################################################################################
In the below code, I wanted to setup the initial tabs to contain all the header information as well as setup some of the formatting for the columns I knew would be standard across 
the sheets.  I took this coding template from a youtube course offered by Trish Connor-Cata at Microsoft.  It successfully ran throughout all my sheets and allowed me to setup the 
columns and headers with the information I needed to see before moving on to the next part of the code.
        
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
        Range("P4").Select
        Selection.NumberFormat = ("#,000")
        Columns("A:P").Select
        Selection.Columns.AutoFit
        Range("A1").Select

'################################################################################################################################################

The below "For" loop was based on an in class activity in order to generate the Ticker Total along with its Total Volume

         For i = 2 To LastRow

             If Cells(i - 1, 1).Value <> Cells(i, 1).Value Then
                 SummaryRowTable = SummaryRowTable + 1
                 Ticker_Name = Cells(i, 1).Value
                 Ticker_Total = Cells(i, 7).Value
                 
                 Range("J" & SummaryRowTable).Value = Ticker_Name
                 Range("M" & SummaryRowTable).Value = Ticker_Total
'################################################################################################################################################

The below code reformats column B into a proper date format.  As the testing data did not come in a date format, I wanted to ensure that I converted this to a date in order to feel confident about applying Min and Max conditions.  Also to ensure once I ran the code against the file they want, that it did not stop due to the cells being different formats. Also, I wanted to also "safe guard" the column to ensure that if the dates were not in the correct order, they would still ensure to grab the true Minimum and Maximum dates instead of simply taking the first line vs the last line. 

                'Refortting of the data column so it could be calculated as a formal date
                 Range("B" & i).Value = DateSerial(Left(Cells(i, 2).Value, 4), Mid(Cells(i, 2).Value, 5, 2), Right(Cells(i, 2).Value, 2))
                 
                 MinDate = Cells(i, 2).Value
                 QOpenBal = Cells(i, 3).Value
          
                 MaxDate = Cells(i, 2).Value
                 QCloseBal = Cells(i, 6).Value
                 
        '################################################################################################################################################
The below code take the reformatted date colums and applies the Boolean expression to ensure the cells we are retrieving are formatted correctly as a date in order to be picked up as a part of the condition. Essentially if the length of the string is still the 8 characters, reapply the correct formatting before moving on to the next part of the condition which is to determine the Min and Max date values.  I received assistance from a co-worker to help me navigate this part of the code.  Most specifically the cStr code.  Working in finance, many of these practices are to ensure we are working with the cleanest data possible before applying calculations. 

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

        '################################################################################################################################################
The below part of the code I was help by a tutor to apply the correct formatting to the cells within an additional For Loop.  I was struggling to not colour the entire column to reduce time, but instead wanted to only format the cells which contain actual data and not "blanks" 

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
                 
                 LastRowL = ws.Cells(ws.Rows.Count, "L").End(xlUp).Row
                 
                 For Each cellL In ws.Range("L2:L" & LastRowL)
                     If IsEmpty(cellL.Value) Then
                         cellL.Interior.ColorIndex = xlNone
                     ElseIf cellL.Value < 0 Then
                         cellL.Interior.Color = RGB(255, 0, 0)
                     Else
                         cellL.Interior.Color = RGB(0, 255, 0)
                     End If
                 
                 Next cellL
                                  
                 Ticker_Total = Ticker_Total + ws.Cells(i, 7).Value
                 Range("M" & SummaryRowTable).Value = Ticker_Total
             End If
     
         Next i

        '################################################################################################################################################
The below code was based on provided help from a classmate - Jose Traboulski.  I was initially struggling with summarizing the data into the correct cells, but once I saw his approach and how he applied the code, I was able to revise my approach and simplify it to work for my code.  Once I really dove deep into the structure, I realized that it was more simple than my original approach. 

         LastRow_Summary = Cells(Rows.Count, 12).End(xlUp).Row
     
         ' Find greatest % increase and respective ticker and place these in the summary table
     
             For j = 3 To LastRow_Summary
                 If Cells(j, 12).Value > MaxValue Then
                     MaxValue = Cells(j, 12).Value
                     Max_Ticker = Cells(j, 10).Value
                 End If
   
         ' Find greatest % decrease and respective ticker and place these in the summary table

                 If Cells(j, 12).Value < MinValue Then
                     MinValue = Cells(j, 12).Value
                     Min_Ticker = Cells(j, 10).Value
                 End If
        
         ' Find greatest total increase and respective ticker and place these in the summary table

                 If Cells(j, 13).Value > MaxVolume Then
                     MaxVolume = Cells(j, 13).Value
                     MaxVolume_Ticker = Cells(j, 10).Value
                 End If
             Next j
         
         'Place values into cells
         Range("O2").Value = Max_Ticker
         Range("P2").Value = MaxValue
         Range("P2").NumberFormat = "0.00%"
         Range("O3").Value = Min_Ticker
         Range("P3").Value = MinValue
         Range("P3").NumberFormat = "0.00%"
         Range("O4").Value = MaxVolume_Ticker
         Range("P4").Value = MaxVolume

        '################################################################################################################################################

Thank you!
