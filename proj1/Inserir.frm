VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Inserir 
   Caption         =   "Produção"
   ClientHeight    =   5835
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   6225
   OleObjectBlob   =   "Inserir.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Inserir"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Coded by @augusto_lopes https://github.com/ALo2021
Private Sub CommandButton1_Click()

Calendario.year.Caption = year(Now)
Calendario.monthh.Caption = monthName(month(Now))
Assigndate
Calendario.Show

End Sub

Private Sub CommandButton2_Click()

    Dim LastRow As Long

    LastRow = Plan2.Range("B" & Rows.Count).End(xlUp).Row + 1

    Plan2.Range("A" & LastRow).Value = ComboBox1.Text
    Plan2.Range("B" & LastRow).Value = DateValue(Label5.Caption)
    Plan2.Range("C" & LastRow).Value = TextBox1.Text
    Plan2.Range("G" & LastRow).Value = ComboBox2.Text
    Plan2.Range("H" & LastRow).Value = TextBox2.Text
    Plan2.Range("I" & LastRow).Value = DateValue(Label7.Caption)
    Plan2.Range("J" & LastRow).Value = ComboBox3.Text
    Plan2.Range("K" & LastRow).Value = TextBox3.Text
    Plan2.Range("L" & LastRow).Value = TextBox4.Text

End Sub

Private Sub CommandButton3_Click()
    Unload Me

End Sub

Private Sub CommandButton4_Click()

Calendario2.year.Caption = year(Now)
Calendario2.monthh.Caption = monthName(month(Now))
Assigndate2
Calendario2.Show

End Sub

Private Sub TextBox_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    If Not IsNumeric(TextBox1.Value) Or IsNumeric(TextBox2.Value) Or IsNumeric(TextBox3.Value) Or IsNumeric(TextBox4.Value) Then
        MsgBox "Apenas valores numericos!!!"
        Cancel = True
    End If

End Sub

Private Sub TextBox2_Change()

End Sub

Private Sub UserForm_Initialize()
ComboBox1.ColumnCount = 2
Dim initial_array(52, 1) As Variant
'populate array
initial_array(0, 0) = Plan1.Range("A5").Value
initial_array(0, 1) = Plan1.Range("B5").Value

initial_array(1, 0) = Plan1.Range("A6").Value
initial_array(1, 1) = Plan1.Range("B6").Value

initial_array(2, 0) = Plan1.Range("A7").Value
initial_array(2, 1) = Plan1.Range("B7").Value

initial_array(3, 0) = Plan1.Range("A8").Value
initial_array(3, 1) = Plan1.Range("B8").Value

initial_array(4, 0) = Plan1.Range("A9").Value
initial_array(4, 1) = Plan1.Range("B9").Value

initial_array(5, 0) = Plan1.Range("A10").Value
initial_array(5, 1) = Plan1.Range("B10").Value

initial_array(6, 0) = Plan1.Range("A11").Value
initial_array(6, 1) = Plan1.Range("B11").Value

initial_array(7, 0) = Plan1.Range("A12").Value
initial_array(7, 1) = Plan1.Range("B12").Value

initial_array(8, 0) = Plan1.Range("A13").Value
initial_array(8, 1) = Plan1.Range("B13").Value

initial_array(9, 0) = Plan1.Range("A14").Value
initial_array(9, 1) = Plan1.Range("B14").Value

initial_array(10, 0) = Plan1.Range("A15").Value
initial_array(10, 1) = Plan1.Range("B15").Value

initial_array(11, 0) = Plan1.Range("A16").Value
initial_array(11, 1) = Plan1.Range("B16").Value

initial_array(12, 0) = Plan1.Range("A17").Value
initial_array(12, 1) = Plan1.Range("B17").Value

initial_array(13, 0) = Plan1.Range("A18").Value
initial_array(13, 1) = Plan1.Range("B18").Value

initial_array(14, 0) = Plan1.Range("A19").Value
initial_array(14, 1) = Plan1.Range("B19").Value

initial_array(15, 0) = Plan1.Range("A20").Value
initial_array(15, 1) = Plan1.Range("B20").Value

initial_array(16, 0) = Plan1.Range("A21").Value
initial_array(16, 1) = Plan1.Range("B21").Value

initial_array(17, 0) = Plan1.Range("A22").Value
initial_array(17, 1) = Plan1.Range("B22").Value

initial_array(18, 0) = Plan1.Range("A23").Value
initial_array(18, 1) = Plan1.Range("B23").Value

initial_array(19, 0) = Plan1.Range("A24").Value
initial_array(19, 1) = Plan1.Range("B24").Value

initial_array(20, 0) = Plan1.Range("A25").Value
initial_array(20, 1) = Plan1.Range("B25").Value

initial_array(21, 0) = Plan1.Range("A26").Value
initial_array(21, 1) = Plan1.Range("B26").Value

initial_array(22, 0) = Plan1.Range("A27").Value
initial_array(22, 1) = Plan1.Range("B27").Value

initial_array(23, 0) = Plan1.Range("A28").Value
initial_array(23, 1) = Plan1.Range("B28").Value

initial_array(24, 0) = Plan1.Range("A29").Value
initial_array(24, 1) = Plan1.Range("B29").Value

initial_array(25, 0) = Plan1.Range("A30").Value
initial_array(25, 1) = Plan1.Range("B30").Value

initial_array(26, 0) = Plan1.Range("A31").Value
initial_array(26, 1) = Plan1.Range("B31").Value

initial_array(27, 0) = Plan1.Range("A32").Value
initial_array(27, 1) = Plan1.Range("B32").Value

initial_array(28, 0) = Plan1.Range("A33").Value
initial_array(28, 1) = Plan1.Range("B33").Value

initial_array(29, 0) = Plan1.Range("A34").Value
initial_array(29, 1) = Plan1.Range("B34").Value

initial_array(30, 0) = Plan1.Range("A35").Value
initial_array(30, 1) = Plan1.Range("B35").Value

initial_array(31, 0) = Plan1.Range("A36").Value
initial_array(31, 1) = Plan1.Range("B36").Value

initial_array(32, 0) = Plan1.Range("A37").Value
initial_array(32, 1) = Plan1.Range("B37").Value

initial_array(33, 0) = Plan1.Range("A38").Value
initial_array(33, 1) = Plan1.Range("B38").Value

initial_array(34, 0) = Plan1.Range("A39").Value
initial_array(34, 1) = Plan1.Range("B39").Value

initial_array(35, 0) = Plan1.Range("A40").Value
initial_array(35, 1) = Plan1.Range("B40").Value

initial_array(36, 0) = Plan1.Range("A41").Value
initial_array(36, 1) = Plan1.Range("B41").Value

initial_array(37, 0) = Plan1.Range("A42").Value
initial_array(37, 1) = Plan1.Range("B42").Value

initial_array(38, 0) = Plan1.Range("A43").Value
initial_array(38, 1) = Plan1.Range("B43").Value

initial_array(39, 0) = Plan1.Range("A44").Value
initial_array(39, 1) = Plan1.Range("B44").Value

initial_array(40, 0) = Plan1.Range("A45").Value
initial_array(40, 1) = Plan1.Range("B45").Value

initial_array(41, 0) = Plan1.Range("A46").Value
initial_array(41, 1) = Plan1.Range("B46").Value

initial_array(42, 0) = Plan1.Range("A47").Value
initial_array(42, 1) = Plan1.Range("B47").Value

initial_array(43, 0) = Plan1.Range("A48").Value
initial_array(43, 1) = Plan1.Range("B48").Value

initial_array(44, 0) = Plan1.Range("A49").Value
initial_array(44, 1) = Plan1.Range("B49").Value

initial_array(45, 0) = Plan1.Range("A50").Value
initial_array(45, 1) = Plan1.Range("B50").Value

initial_array(46, 0) = Plan1.Range("A51").Value
initial_array(46, 1) = Plan1.Range("B51").Value

initial_array(47, 0) = Plan1.Range("A52").Value
initial_array(47, 1) = Plan1.Range("B52").Value

initial_array(48, 0) = Plan1.Range("A53").Value
initial_array(48, 1) = Plan1.Range("B53").Value

initial_array(49, 0) = Plan1.Range("A54").Value
initial_array(49, 1) = Plan1.Range("B54").Value

initial_array(50, 0) = Plan1.Range("A55").Value
initial_array(50, 1) = Plan1.Range("B55").Value

initial_array(51, 0) = Plan1.Range("A56").Value
initial_array(51, 1) = Plan1.Range("B56").Value

initial_array(52, 0) = Plan1.Range("A57").Value
initial_array(52, 1) = Plan1.Range("B57").Value

'then populate combobox with full array
ComboBox1.List = initial_array
ComboBox1.ColumnWidths = "18"

'segunda combobox
Dim array_ini(2) As Variant

array_ini(0) = Plan1.Range("W43").Value
array_ini(1) = Plan1.Range("W44").Value
array_ini(2) = Plan1.Range("W45").Value

ComboBox2.List = array_ini

'Operadores
Dim ope_ini(12) As Variant

ope_ini(0) = Plan1.Range("W47").Value
ope_ini(1) = Plan1.Range("W48").Value
ope_ini(2) = Plan1.Range("W49").Value
ope_ini(3) = Plan1.Range("W50").Value
ope_ini(4) = Plan1.Range("W51").Value
ope_ini(5) = Plan1.Range("W52").Value
ope_ini(6) = Plan1.Range("W53").Value
ope_ini(7) = Plan1.Range("W54").Value
ope_ini(8) = Plan1.Range("W55").Value
ope_ini(9) = Plan1.Range("W56").Value
ope_ini(10) = Plan1.Range("W57").Value
ope_ini(11) = Plan1.Range("W58").Value

ComboBox3.List = ope_ini

End Sub
