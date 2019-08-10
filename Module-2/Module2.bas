Attribute VB_Name = "Module1"
Sub Mod_Homework()

For Each ws In Worksheets

'Establish Variables
Dim Ticker As String
Dim YearBeg As Double
Dim YearEnd As Double
Dim Volume As Double
Dim Summary_Table_Row As Integer

'Create Summary Table Headers
Summary_Table_Row = 2
ws.Range("I1").Value = "Ticker"
ws.Range("J1").Value = "Yearly Price Change"
ws.Range("K1").Value = "Percent Change"
ws.Range("L1").Value = "Total Volume"
LastRow = ws.Cells(Rows.Count, 1).End(xlUp).Row

For r = 2 To LastRow

'Find and Save your Opening Balance
If ws.Cells(r, 1).Value <> ws.Cells(r - 1, 1).Value Then
    YearBeg = ws.Cells(r, 3).Value
End If

    'If the current ticker does not match the next ticker..
    If ws.Cells(r + 1, 1).Value <> ws.Cells(r, 1).Value Then
        
        'Assigns ticker name to be shown in the summary
        Ticker = ws.Cells(r, 1).Value
        
        'Find and Save your Closing Balance
        YearEnd = ws.Cells(r, 6).Value
        
        'Determine Final Volume
        Volume = Volume + ws.Cells(r, 7).Value
                
        'Print Tickers in Summary Line
        ws.Range("I" & Summary_Table_Row).Value = Ticker
        
        'Calculate and Print Dollar & Percentage Changes
        YearChg = YearEnd - YearBeg
        ws.Range("J" & Summary_Table_Row).Value = YearChg
                        
            'Conditional Formatting for Yearly Change
            If YearChg >= 0 Then
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 4
            Else
                ws.Range("J" & Summary_Table_Row).Interior.ColorIndex = 3
            End If
            
        'Calculate, print, and format Yearly % Change
        If YearBeg <> 0 Then
            ws.Range("K" & Summary_Table_Row).Value = YearChg / YearBeg
        Else
            ws.Range("K" & Summary_Table_Row).Value = "0"
        End If
        ws.Range("K" & Summary_Table_Row).Style = "Percent"
        
        'Print volume
        ws.Range("L" & Summary_Table_Row).Value = Volume
        
        'Adjust/Reset the Variables
        Summary_Table_Row = Summary_Table_Row + 1
        YearBeg = 0
        Volume = 0
        
    'Or if the ticker is the same...
    Else
        '...Add the volume to the running total
        Volume = Volume + ws.Cells(r, 7).Value
    End If
Next r



'Hard Portion
'Establish the variables you're looking for
Dim Max As Double
Dim Min As Double

'Find the Max %, Print and Label it
Max = WorksheetFunction.Max(Range("K2:K500"))
ws.Range("O2").Value = Max
ws.Range("N2").Value = "Greatest % Increase"

'Find the Min %, Print and Label it
Min = WorksheetFunction.Min(Range("K2:K500"))
ws.Range("O3").Value = Min
ws.Range("N3").Value = "Greatest % Decrease"

'Format the two figures as %
ws.Range("O2:O3").Style = "Percent"

'Find the Max Volume, Print and Label it
Max = WorksheetFunction.Max(Range("L2:L500"))
ws.Range("O4").Value = Max
ws.Range("N4").Value = "Greatest Total Volume"

Next ws


End Sub
''Questions-- Can 'LastRow' not be used in the Max/Min functions?
''Formatting multiple decimal places
''Does the sub pull information from the top sheet when 'ws' isn't used
