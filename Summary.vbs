Sub totals()

'get greatest increase
 
        Cells(2, 17).Value = WorksheetFunction.Max(Range("L2:L3169"))
        
        For r = 2 To 3169
        If Cells(r, 12).Value = Cells(2, 17) Then
        Cells(2, 16).Value = Cells(r, 10).Value
        End If
   
'get greatest decrease

        Cells(3, 17).Value = WorksheetFunction.Min(Range("L2:L3169"))
        
        If Cells(r, 12) = Cells(3, 17) Then
        Cells(3, 16).Value = Cells(r, 10).Value
        End If
        

'Get greatest volume
        
        Cells(4, 17).Value = WorksheetFunction.Max(Range("M2:M3169"))
       
        If Cells(r, 13) = Cells(4, 17) Then
        Cells(4, 16).Value = Cells(r, 10).Value
        End If
    Next r

Cells(2, 15).Value = "Greatest % increase"
Cells(3, 15).Value = "Greatest % decrease"
Cells(4, 15).Value = "Greatest volume"
Cells(1, 16).Value = "Ticker"
Cells(1, 17).Value = "Value"

End Sub
