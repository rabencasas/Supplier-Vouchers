Attribute VB_Name = "Module2"
Sub Button16_Click()
    ' total deduction
    Sheets("form").Range("R14").Value = Sheets("form").Range("R10").Value + Sheets("form").Range("R11").Value + Sheets("form").Range("R12").Value
    ' net amount
    Sheets("form").Range("R16").Value = Sheets("form").Range("F11").Value - Sheets("form").Range("R14").Value
End Sub
Sub Button8_Click()
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
    
    Sheets("Disbursement_Voucher").Range("B64").Value = "Document prepared by: " & Sheets("Configuration").Range("D4").Value
    
    Sheets("Disbursement_Voucher").PrintPreview
End Sub
Sub Button9_Click()
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
    If Sheets("Configuration").Range("D12").Value = "Print Voucher first" Then
        Sheets("Disbursement_Voucher").PrintOut From:=1, To:=1, Copies:=Sheets("Configuration").Range("D14").Value
        Sheets("alobs").PrintOut From:=1, To:=1, Copies:=Sheets("Configuration").Range("D16").Value
    Else
        Sheets("alobs").PrintOut From:=1, To:=1, Copies:=Sheets("Configuration").Range("D16").Value
        Sheets("Disbursement_Voucher").PrintOut From:=1, To:=1, Copies:=Sheets("Configuration").Range("D14").Value
    End If
End Sub
Sub Button10_Click()
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
    Sheets("Disbursement_Voucher").Range("G23").Value = Sheets("form").Range("P10").Value
    Sheets("Disbursement_Voucher").Range("G24").Value = Sheets("form").Range("P11").Value
    Sheets("Disbursement_Voucher").Range("G25").Value = Sheets("form").Range("P12").Value
    ' Deduction Amount
    Sheets("Disbursement_Voucher").Range("H22").Value = Sheets("form").Range("R10").Value
    Sheets("Disbursement_Voucher").Range("H23").Value = Sheets("form").Range("R11").Value
    Sheets("Disbursement_Voucher").Range("H24").Value = Sheets("form").Range("R12").Value
    ' Total Deduction & Net Amount
    Sheets("Disbursement_Voucher").Range("L24").Value = Sheets("form").Range("R14").Value
    Sheets("Disbursement_Voucher").Range("L26").Value = Sheets("form").Range("R16").Value
    ' Footer
    
    Sheets("Disbursement_Voucher").Range("B63").Value = "Document prepared by: " & Sheets("Configuration").Range("D4").Value
    
    If Sheets("Configuration").Range("D8").Value = "" Then
        MsgBox "Log Directory is not set on 'Configuration' sheet.", vbOKCancel, "Log Directory not set"
    Else
         
        On Error GoTo errmsg
         
        Dim file As String
        Dim textfile As Integer
        
        file = Sheets("Configuration").Range("D8").Value & "\" & Sheets("Form").ComboBox1.Text & " - [" & Format(Sheets("Form").Range("f11").Value, "standard") & "].voucher"
        
        textfile = FreeFile
        
        Open file For Output As textfile
        
        Print #textfile, Sheets("form").ComboBox1.Text
        Print #textfile, Sheets("form").Range("F7").Value
        Print #textfile, Sheets("form").Range("F9").Value & vbNewLine
        
        Print #textfile, "Amount (Gross): " & Format(Sheets("form").Range("F11").Value, "standard")
        Print #textfile, "Service Amount: " & Format(Sheets("form").Range("F18").Value, "standard")
        Print #textfile, "Goods Amount: " & Format(Sheets("form").Range("F20").Value, "standard")
        Print #textfile, Sheets("form").Range("F13").Value & vbNewLine
        
        Print #textfile, "Deductions:"
        Print #textfile, vbTab & Sheets("form").Range("P10").Value & ":" & vbTab & Format(Sheets("form").Range("R10").Value, "standard")
        Print #textfile, vbTab & Sheets("form").Range("P11").Value & ":" & vbTab & Format(Sheets("form").Range("R11").Value, "standard")
        If Sheets("form").Range("F9").Value = "VAT - Goods & Services" Or Sheets("form").Range("F9").Value = "NON-VAT - Goods & Services" Then
            Print #textfile, vbTab & Sheets("form").Range("P12").Value & ":" & vbTab & Sheets("form").Range("R12").Value
        End If
        Print #textfile, vbTab & "Total: " & Format(Sheets("form").Range("R14").Value, "standard") & vbNewLine
        
        Print #textfile, "Net Amount: " & Format(Sheets("form").Range("R16").Value, "standard") & vbNewLine
        
        Print #textfile, "Purpose of purcahse:"
        Print #textfile, vbTab & Sheets("form").Range("F22").Value & vbNewLine
        
        Print #textfile, "Requesting Personnel: " & Sheets("form").Range("T5").Value
        Print #textfile, "Clerk: " & Sheets("Configuration").Range("D4").Value & vbNewLine
        
        Print #textfile, Format(Now, "hh:mm am/pm" & vbNewLine & "mmm-dd-yyyy")
        
        Close textfile
    
        MsgBox "Voucher successfully logged!" & vbNewLine & vbNewLine & Sheets("form").ComboBox1.Text & vbNewLine & "Amount: " & Format(Sheets("form").Range("F11").Value, "standard") & vbNewLine, vbOKOnly
        
errmsg:
        If Err.Number > 0 Then
            MsgBox "An unexpectederror has occured!" & vbNewLine & "Please check the voucher log directory path in the 'Configuration' sheet", vbCritical, "Unexpected error occured"
        End If
    End If
End Sub
