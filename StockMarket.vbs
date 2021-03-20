Option Explicit

'define variables

    Dim closing_value As Double
    Dim opening_value As Double
    Dim change As Double
    Dim percentchange As Double
    Dim volume As LongLong
    Dim rw As Long 'row number
    Dim t As Integer 'ticker row
    Dim lastrow As Long

lastrow = Cells(Rows.Count, "A").End(xlUp).Row

'Get Unique names
    ActiveSheet.Range("A1:A" & lastrow).AdvancedFilter _
    Action:=xlFilterCopy, _
    CopyToRange:=ActiveSheet.Range("J1"), _
    Unique:=True

'get closing value
           For t = 2 To 4000
                For rw = 2 To lastrow
                If Cells(rw, 1).Value = Cells(t, 10).Value Then
                closing_value = Cells(rw, 6)
                Cells(t, 15) = closing_value

                End If
                Next rw

        Next t

'get open value
        Dim first As Integer
        first = 0
        For t = 2 To 4000
            If first = 0 Then
                For rw = 2 To lastrow
                If (Cells(rw, 1).Value = Cells(t, 10).Value) And (first = 0) Then
                opening_value = Cells(rw, 3)
                first = first + 1
                Cells(t, 16) = opening_value
            End If
            Next rw
        End If
        first = 0

    Next t

'get change
    For t = 2 To 4000
    Cells(t, 11) = (Cells(t, 15).Value - Cells(t, 16).Value)
    Next t
'get percent change
        For t = 2 To 4000
        If cells(t,11). value <>0 and cells(t,16).value <>0 then
        Cells(t, 12) = FormatPercent(Cells(t, 11).Value / Cells(t, 16).Value)
        End if
        Next t

'get volume

    For t = 2 To 4000
        volume = Application.WorksheetFunction.SumIf(Range("A2:A" & lastrow), Cells(t, 10).Value, Range("G2:G" & lastrow))
        Cells(t, 13) = volume
    Next t
    For r = 2 To 4000
        If Cells(r, 11).Value > 0 Then
            Cells(r, 11).Interior.ColorIndex = 4
            Else
                Cells(r, 11).Interior.ColorIndex = 3
             End If
        Next r
    Cells(1, 10) = "Ticker"
    Cells(1, 11) = "Change"
    Cells(1, 12) = "Percent Change"
    Cells(1, 13) = "Total Volume"
  

    Range("O2:P4000").Clear

End Sub


