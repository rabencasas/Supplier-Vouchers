Attribute VB_Name = "Module1"
Sub Button1_Click()
    '*** GOODS ===================================================================================================================

    'If VAT-Goods Only
    If Sheets("form").Range("F9").Value = "VAT - Goods Only" Then
        'Clear Service & Labor Amount
        Sheets("form").Range("F18").Value = "N/A"
        Sheets("form").Range("F20").Value = "N/A"
    
        'Set deduction labels
        Sheets("form").Range("P10").Value = "5% VAT"
        Sheets("form").Range("P11").Value = "1% wt"
        Sheets("form").Range("P12").Value = ""
        
        'vat 5%
        Sheets("form").Range("U10").Value = Sheets("form").Range("F11").Value / 1.12 * 0.05
        Sheets("form").Range("R10").Value = Format(Sheets("form").Range("F11").Value / 1.12 * 0.05, "standard")
        'wt 1%
        Sheets("form").Range("U11").Value = Sheets("form").Range("F11").Value / 1.12 * 0.01
        Sheets("form").Range("R11").Value = Format(Sheets("form").Range("F11").Value / 1.12 * 0.01, "standard")
        'clear vacant wt
        Sheets("form").Range("R12").Value = ""
        Sheets("form").Range("U12").Value = ""
    End If
    
    'If NON-VAT-Goods Only
    If Sheets("form").Range("F9").Value = "NON-VAT - Goods Only" Then
        'Clear Service & Labor Amount
        Sheets("form").Range("F18").Value = "N/A"
        Sheets("form").Range("F20").Value = "N/A"
    
        'Set deduction labels
        Sheets("form").Range("P10").Value = "3% NV"
        Sheets("form").Range("P11").Value = "1% wt"
        Sheets("form").Range("P12").Value = ""
        
        'vat 3%
        Sheets("form").Range("U10").Value = Sheets("form").Range("F11").Value * 0.03
        Sheets("form").Range("R10").Value = Format(Sheets("form").Range("F11").Value * 0.03, "standard")
        'wt 1%
        Sheets("form").Range("U11").Value = Sheets("form").Range("F11").Value * 0.01
        Sheets("form").Range("R11").Value = Format(Sheets("form").Range("F11").Value * 0.01, "standard")
        'clear vacant wt
        Sheets("form").Range("R12").Value = ""
        Sheets("form").Range("U12").Value = ""
    End If
    
    '*** SERVICES ===================================================================================================================

    'If VAT-Services Only
    If Sheets("form").Range("F9").Value = "VAT - Services Only" Then
        'Clear Service & Labor Amount
        Sheets("form").Range("F18").Value = "N/A"
        Sheets("form").Range("F20").Value = "N/A"
    
        'Set deduction labels
        Sheets("form").Range("P10").Value = "5% VAT"
        Sheets("form").Range("P11").Value = "2% wt"
        Sheets("form").Range("P12").Value = ""
        
        'vat 5%
        Sheets("form").Range("U10").Value = Sheets("form").Range("F11").Value / 1.12 * 0.05
        Sheets("form").Range("R10").Value = Format(Sheets("form").Range("F11").Value / 1.12 * 0.05, "standard")
        'wt 1%
        Sheets("form").Range("U11").Value = Sheets("form").Range("F11").Value / 1.12 * 0.02
        Sheets("form").Range("R11").Value = Format(Sheets("form").Range("F11").Value / 1.12 * 0.02, "standard")
        'clear vacant wt
        Sheets("form").Range("R12").Value = ""
        Sheets("form").Range("U12").Value = ""
    End If
    
    'If NON-VAT-Services Only
    If Sheets("form").Range("F9").Value = "NON-VAT - Services Only" Then
        'Clear Service & Labor Amount
        Sheets("form").Range("F18").Value = "N/A"
        Sheets("form").Range("F20").Value = "N/A"
    
        'Set deduction labels
        Sheets("form").Range("P10").Value = "3% NV"
        Sheets("form").Range("P11").Value = "2% wt"
        Sheets("form").Range("P12").Value = ""
        
        'vat 3%
        Sheets("form").Range("U10").Value = Sheets("form").Range("F11").Value * 0.03
        Sheets("form").Range("R10").Value = Format(Sheets("form").Range("F11").Value * 0.03, "standard")
        'wt 1%
        Sheets("form").Range("U11").Value = Sheets("form").Range("F11").Value * 0.02
        Sheets("form").Range("R11").Value = Format(Sheets("form").Range("F11").Value * 0.02, "standard")
        'clear vacant wt
        Sheets("form").Range("R12").Value = ""
        Sheets("form").Range("U12").Value = ""
    End If
    
    '*** GOODS & SERVICES ===========================================================================================

    'If VAT-Goods & Services
    If Sheets("form").Range("F9").Value = "VAT - Goods & Services" Then
        If Sheets("form").Range("F18").Value <> "N/A" And Sheets("form").Range("F20").Value <> "N/A" Then
            'Set deduction labels
            Sheets("form").Range("P10").Value = "5% VAT"
            Sheets("form").Range("P11").Value = "2% wt"
            Sheets("form").Range("P12").Value = "1% wt"
            
            'vat 5%
            Sheets("form").Range("U10").Value = Sheets("form").Range("F11").Value / 1.12 * 0.05
            Sheets("form").Range("R10").Value = Format(Sheets("form").Range("F11").Value / 1.12 * 0.05, "standard")
            'wt 2%
            Sheets("form").Range("U11").Value = Sheets("form").Range("F18").Value / 1.12 * 0.02
            Sheets("form").Range("R11").Value = Format(Sheets("form").Range("F18").Value / 1.12 * 0.02, "standard")
            'wt 1%
            Sheets("form").Range("U12").Value = Sheets("form").Range("F20").Value / 1.12 * 0.01
            Sheets("form").Range("R12").Value = Format(Sheets("form").Range("F20").Value / 1.12 * 0.01, "standard")
        Else
            MsgBox "Please input the 'Service Amount (2%)' and the 'Goods Amount (1%)' to continue.", vbCritical
        End If
    End If
    
    'If NON-VAT-Goods & Services
    If Sheets("form").Range("F9").Value = "NON-VAT - Goods & Services" Then
        If Sheets("form").Range("F18").Value <> "N/A" And Sheets("form").Range("F20").Value <> "N/A" Then
            'Set deduction labels
            Sheets("form").Range("P10").Value = "3% NV"
            Sheets("form").Range("P11").Value = "2% wt"
            Sheets("form").Range("P12").Value = "1% wt"
            
            'vat 3%
            Sheets("form").Range("U10").Value = Sheets("form").Range("F11").Value * 0.03
            Sheets("form").Range("R10").Value = Format(Sheets("form").Range("F11").Value * 0.03, "standard")
            'wt 2%
            Sheets("form").Range("U11").Value = Sheets("form").Range("F18").Value * 0.02
            Sheets("form").Range("R11").Value = Format(Sheets("form").Range("F18").Value * 0.02, "standard")
            'wt 1%
            Sheets("form").Range("U12").Value = Sheets("form").Range("F20").Value * 0.01
            Sheets("form").Range("R12").Value = Format(Sheets("form").Range("F20").Value * 0.01, "standard")
        Else
            MsgBox "Please input the 'Service Amount (2%)' and the 'Goods Amount (1%)' to continue.", vbCritical
        End If
    End If
    
        ' total deduction
         Sheets("form").Range("R14").Value = Sheets("form").Range("R10").Value + Sheets("form").Range("R11").Value + Sheets("form").Range("R12").Value
        ' net amount
        Sheets("form").Range("R16").Value = Sheets("form").Range("F11").Value - Sheets("form").Range("R14").Value
End Sub
