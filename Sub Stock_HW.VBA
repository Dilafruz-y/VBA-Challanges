
Sub Stock()
'* Create a script that will loop through all the stocks for one year and output the following information.
Dim CWS As Worksheet

'Loop Through Worksheets
For Each CWS In Worksheets
 
  
  ' Set an initial variable for holding the tricker
  

  ' Set an initial variable for Vol Total
  Dim Vol_Total As Double
  Vol_Total = 0

    ' Set an initial variable for Change
  Dim Change_Total As Double
  Change_Total = 0

    Dim Persent_Change As Double
    Persent_Change = 0
'
    
  ' Keep track of the lcation for each tricker in the summary table
  Dim Summary_Table_Row As Integer
  Summary_Table_Row = 2
    ' Determine the Last Row
        Lastrow = CWS.Cells(Rows.Count, 1).End(xlUp).Row

Dim start_row As Long
start_row = 2

  ' Loop through Trickers
  For i = 2 To Lastrow

    ' Check if we are still within the same Tricker, if it is not...
    If CWS.Cells(i + 1, 1).Value <> CWS.Cells(i, 1).Value Then

      ' Set the Tricker
      Ticker = Cells(i, 1).Value

      ' Add to the Vol Total
      Vol_Total = Vol_Total + CWS.Cells(i, 7).Value
     
     'Calculate change
     
     Change_Total = CWS.Cells(i, 6) - CWS.Cells(start_row, 3)
    
    
    'Calculate Persent_change
    
    
     If CWS.Cells(start_row, 3) <> 0 Then
     
     Persent_Change = Round((Change_Total / CWS.Cells(start_row, 3) * 100), 2)
    
     End If
      
      ' Print the Ticker in the Summary Table
      CWS.Range("K" & Summary_Table_Row).Value = Ticker

      
       
      'Print the Change Total in the Summary Table
      CWS.Range("L" & Summary_Table_Row).Value = Change_Total
      
       'Print the Persantage Change in the Summary Table
      CWS.Range("M" & Summary_Table_Row).Value = (CStr(Persent_Change) & "%")
      
      ' Print the Vol Total to the Summary Table
      CWS.Range("N" & Summary_Table_Row).Value = Vol_Total
      
      
      ' Add one to the summary table row
      Summary_Table_Row = Summary_Table_Row + 1
      
      ' Reset the Vol Total
    Vol_Total = 0
    Change_Total = 0
    Persent_Change = 0

    ' If the cell immediately following a row is the same tricker...
    start_row = i + 1
    Else

      ' Add to the Vol Total
      Vol_Total = Vol_Total + Cells(i, 7).Value
   
      
     End If

  Next i
 'Add header the Vol Total
      CWS.Cells(1, 11) = " Ticker "
      CWS.Cells(1, 12) = "Change_Total"
      CWS.Cells(1, 13) = "Persent_Change"
      CWS.Cells(1, 14) = "Total_Vol"
      
     'Determine the Last Row
        Lastrow = CWS.Cells(Rows.Count, 1).End(xlUp).Row

    For i = 2 To Lastrow

        If CWS.Cells(i, 12).Value > 0 Then
    CWS.Cells(i, 12).Interior.ColorIndex = 4
    Else
    End If
    
    If CWS.Cells(i, 12).Value < 0 Then
    CWS.Cells(i, 12).Interior.ColorIndex = 3
    Else
        End If
    Next i
Next CWS

End Sub
