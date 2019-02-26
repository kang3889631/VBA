Sub Button15_Click()

Dim value, i, count_hihi, count_hi, count_low, count_lowlow, limit_hihi, limit_hi, limit_low, limit_lowlow, delay_hihi, delay_hi, delay_low, delay_lowlow, alarm_hi, i_hihi As Integer

alarm_hihi = 0
alarm_hi = 0
alarm_lowlow = 0
alarm_low = 0

delay_hihi = Cells(12, 11).value
delay_hi = Cells(13, 11).value
delay_low = Cells(14, 11).value
delay_lowlow = Cells(15, 11).value

limit_hihi = Cells(12, 10).value
limit_hi = Cells(13, 10).value
limit_low = Cells(14, 10).value
limit_lowlow = Cells(15, 10).value
count_hihi = 0
count_hi = 0
count_low = 0
count_lowlow = 0

i_hihi = 22
i_hi = 22
i_low = 22
i_lowlow = 22

For i = 19 To 86419
    value = Cells(i, 5).value
    
    'Hihi alarm
    If (value >= limit_hihi) Then 'over the hihi limit, count one
    count_hihi = count_hihi + 1
        
    ElseIf (value < limit_hihi) Then
        If (count_hihi >= delay_hihi) Then 'whenever reach to the delay, count one for alarm, and refresh count_hihi
            alarm_hihi = alarm_hihi + 1
            Cells(i_hihi, 14).value = count_hihi
            i_hihi = i_hihi + 1
        End If
        count_hihi = 0
    End If
    
    'Hi alarm
    If (value >= limit_hi) Then 'over the hi limit, count one
    count_hi = count_hi + 1
        
    ElseIf (value < limit_hi) Then
        If (count_hi >= delay_hi) Then 'whenever reach to the delay, count one for alarm, and refresh count_hi
            alarm_hi = alarm_hi + 1
            Cells(i_hi, 15).value = count_hi
            i_hi = i_hi + 1
        End If
        count_hi = 0
    End If
    
    'Low alarm
    If (value >= limit_low) Then 'over the low limit, count one
    count_low = count_low + 1
        
    ElseIf (value < limit_low) Then
        If (count_low >= delay_low) Then 'whenever reach to the delay, count one for alarm, and refresh count_low
            alarm_low = alarm_low + 1
            Cells(i_low, 16).value = count_low
            i_low = i_low + 1
        End If
        count_low = 0
    End If


    'Lowlow alarm
    If (value >= limit_lowlow) Then 'over the lowlow limit, count one
    count_lowlow = count_lowlow + 1
        
    ElseIf (value < limit_lowlow) Then
        If (count_lowlow >= delay_lowlow) Then 'whenever reach to the delay, count one for alarm, and refresh count_lowlow
            alarm_lowlow = alarm_lowlow + 1
            Cells(i_lowlow, 17).value = count_lowlow
            i_lowlow = i_lowlow + 1
        End If
        count_lowlow = 0
    End If
    
Next i
    
    Cells(21, 14).value = alarm_hihi
    Cells(21, 15).value = alarm_hi
    Cells(21, 16).value = alarm_low
    Cells(21, 17).value = alarm_lowlow
    

End Sub

