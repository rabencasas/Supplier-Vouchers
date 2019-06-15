VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True

Private Sub cmdFixForm_Click()
    Range("F7").Value = "=VLOOKUP(E91,C93:N3000,7,FALSE)"
    Range("F13").Value = "=SpellNumber(F11)"
    Range("R8").Value = "=F11"
    Range("O4").Value = "=VLOOKUP(E91,C93:N300,12,FALSE)"
    
    Range("G35").Value = "=Configuration!D4"
    Range("G36").Value = "=Configuration!D8"
    Range("G37").Value = "=Configuration!D12"
    Range("L36").Value = "=Configuration!D14"
    Range("L37").Value = "=Configuration!D16"
    
    Sheets("ALOBS").Range("B25").Value = "=VLOOKUP(FORM!T91,FORM!R93:T128,3,FALSE)"
    
    MsgBox "Done Fix!", vbInformation
End Sub

Private Sub ComboBox1_GotFocus()
    ComboBox1.SelStart = 0
    ComboBox1.SelLength = Me.ComboBox1.TextLength
End Sub

Private Sub ComboBox1_KeyUp(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = 13 Or KeyCode = 9 Or KeyCode = 39 Then
       Range("F7").Activate
   End If
End Sub
