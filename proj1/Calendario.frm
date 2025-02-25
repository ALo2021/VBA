VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Calendario 
   Caption         =   "Calendario"
   ClientHeight    =   5685
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6900
   OleObjectBlob   =   "Calendario.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Calendario"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'https://github.com/ALo2021
Private Sub CommandButton1_Click()

Calendario.year.Caption = Calendario.year - 1
Assigndate

End Sub

Private Sub CommandButton2_Click()

Calendario.year.Caption = Calendario.year + 1
Assigndate


End Sub

Private Sub CommandButton3_Click()

If Calendario.monthh.Caption = "dezembro" Then
Calendario.monthh.Caption = "janeiro"
Else
Calendario.monthh.Caption = monthName(month(DateValue("1 " & Calendario.monthh & " 2024")) + 1)
End If
Assigndate

End Sub

'Coded by @augusto_lopes https://github.com/ALo2021
Private Sub CommandButton4_Click()
If Calendario.monthh.Caption = "janeiro" Then
Calendario.monthh.Caption = "dezembro"
Else
Calendario.monthh.Caption = monthName(month(DateValue("1 " & Calendario.monthh & " 2024")) - 1)
End If
Assigndate

End Sub

Private Sub Frame1_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame10_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame11_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame12_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame13_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame14_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame15_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame16_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame17_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame18_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame19_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame2_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame20_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame21_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame22_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame23_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame24_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame25_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame26_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame27_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame28_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame29_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame3_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame30_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame31_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame32_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame33_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame34_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame35_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame36_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame37_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame38_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame39_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame4_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame40_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame41_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame42_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame5_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame6_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame7_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame8_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub Frame9_Click()
Inserir.Label5.Caption = Calendario.Controls(ActiveControl.Name).Caption & "/" & MonthNameToNumber(Calendario.monthh.Caption) & "/" & Calendario.year.Caption
Unload Me
End Sub

Private Sub UserForm_Click()

End Sub

Function MonthNameToNumber(monthName As String) As Integer
    Select Case LCase(monthName)
        Case "janeiro": MonthNameToNumber = 1
        Case "fevereiro": MonthNameToNumber = 2
        Case "março": MonthNameToNumber = 3
        Case "abril": MonthNameToNumber = 4
        Case "maio": MonthNameToNumber = 5
        Case "junho": MonthNameToNumber = 6
        Case "julho": MonthNameToNumber = 7
        Case "agosto": MonthNameToNumber = 8
        Case "setembro": MonthNameToNumber = 9
        Case "outubro": MonthNameToNumber = 10
        Case "novembro": MonthNameToNumber = 11
        Case "dezembro": MonthNameToNumber = 12
        Case Else: MonthNameToNumber = 0 ' Invalid month name
    End Select
End Function
