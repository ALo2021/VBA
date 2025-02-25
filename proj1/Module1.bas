Attribute VB_Name = "Module1"
'Coded by @augusto_lopes https://github.com/ALo2021
Sub Assigndate()

a = DateValue("1 " & Calendario.monthh & " " & Calendario.year)
primeiroDom = a - Day(a) - Weekday(a - Day(a), vbSunday) + 8


If Day(primeiroDom + 1) <= 2 Then
primeiroDia = primeiroDom - 1
Else
primeiroDia = primeiroDom - 8
End If

Dim Frame As Control
For Each Frame In Calendario.Controls
If TypeName(Frame) = "Frame" Then
i = i + 1

Frame.Caption = Format(Day(primeiroDia + i), "0#")
If monthName(month(primeiroDia + i)) = Calendario.monthh Then
Frame.Enabled = True
Else
Frame.Enabled = False
End If
End If
Next


End Sub
