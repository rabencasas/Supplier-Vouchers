Attribute VB_Name = "Module4"
Sub Button28_Click()
    '*** ACTUAL PRINT ***'
    'Supplier Name & Address
    Sheets("Disbursement_Voucher").Range("D10").Value = Sheets("form").ComboBox1.Text
    Sheets("Disbursement_Voucher").Range("D12").Value = Sheets("form").Range("F7").Value
    
    'AMOUNT, Deduction, Purpose of Purchase & Net Amount
    ' Amount
    Sheets("Disbursement_Voucher").Range("B18").Value = Sheets("form").Range("F13").Value
    Sheets("Disbursement_Voucher").Range("L18").Value = Sheets("form").Range("F11").Value
    ' Purpose of purchase
    Sheets("Disbursement_Voucher").Range("C19").Value = Sheets("form").Range("F22").Value
    ' Deduction Labels
    Sheets("Disbursement_Voucher").Range("G22").Value = Sheets("form").Range("P10").Value
    Sheets("Disbursement_Voucher").Range("G23").Value = Sheets("form").Range("P11").Value
    Sheets("Disbursement_Voucher").Range("G24").Value = Sheets("form").Range("P12").Value
    ' Deduction Amount
    Sheets("Disbursement_Voucher").Range("H22").Value = Sheets("form").Range("R10").Value
    Sheets("Disbursement_Voucher").Range("H23").Value = Sheets("form").Range("R11").Value
    Sheets("Disbursement_Voucher").Range("H24").Value = Sheets("form").Range("R12").Value
    ' Total Deduction & Net Amount
    Sheets("Disbursement_Voucher").Range("L24").Value = Sheets("form").Range("R14").Value
    Sheets("Disbursement_Voucher").Range("L26").Value = Sheets("form").Range("R16").Value
    ' Footer
    
    Sheets("Disbursement_Voucher").Range("B63").Value = "Document prepared by: " & Sheets("Configuration").Range("D4").Value
    
    ' PRINT SEQUENCE
    Sheets("Disbursement_Voucher").PrintOut From:=1, To:=1, Copies:=Sheets("FORM").Range("R26").Value
End Sub
Sub Button29_Click()
    '*** ACTUAL PRINT ***'
    'Supplier Name & Address
    Sheets("Disbursement_Voucher").Range("D10").Value = Sheets("form").ComboBox1.Text
    Sheets("Disbursement_Voucher").Range("D12").Value = Sheets("form").Range("F7").Value
    
    'AMOUNT, Deduction, Purpose of Purchase & Net Amount
    ' Amount
    Sheets("Disbursement_Voucher").Range("B18").Value = Sheets("form").Range("F13").Value
    Sheets("Disbursement_Voucher").Range("L18").Value = Sheets("form").Range("F11").Value
    ' Purpose of purchase
    Sheets("Disbursement_Voucher").Range("C19").Value = Sheets("form").Range("F22").Value
    ' Deduction Labels
    Sheets("Disbursement_Voucher").Range("G22").Value = Sheets("form").Range("P10").Value
    Sheets("Disbursement_Voucher").Range("G23").Value = Sheets("form").Range("P11").Value
    Sheets("Disbursement_Voucher").Range("G24").Value = Sheets("form").Range("P12").Value
    ' Deduction Amount
    Sheets("Disbursement_Voucher").Range("H23").Value = Sheets("form").Range("R10").Value
    Sheets("Disbursement_Voucher").Range("H24").Value = Sheets("form").Range("R11").Value
    Sheets("Disbursement_Voucher").Range("H25").Value = Sheets("form").Range("R12").Value
    ' Total Deduction & Net Amount
    Sheets("Disbursement_Voucher").Range("L24").Value = Sheets("form").Range("R14").Value
    Sheets("Disbursement_Voucher").Range("L26").Value = Sheets("form").Range("R16").Value
    ' Footer
    
    Sheets("Disbursement_Voucher").Range("B63").Value = "Document prepared by: " & Sheets("Configuration").Range("D4").Value
    
    ' PRINT SEQUENCE
    Sheets("ALOBS").PrintOut From:=1, To:=1, Copies:=Sheets("FORM").Range("R27").Value
End Sub
Sub Button30_Click()
    Sheets("form").Range("F7").Value = "=VLOOKUP(E91,C93:N160,7,FALSE)"
End Sub
