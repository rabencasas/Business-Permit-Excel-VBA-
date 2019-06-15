Attribute VB_Name = "mod_XclApp"
Public empNo, capitalInvest As Double

Public Function PrepareDatas()
    'Public permitType, dateIssued, permitNo, bsnOwner, bsnName, bsnAddr, bsnType, orNumber, orDate As Range
    
    Set Sheet4.permitType = Sheet4.Range("H5")
    Set Sheet4.dateIssued = Sheet4.Range("H7")
    Set Sheet4.permitNo = Sheet4.Range("H9")
    Set Sheet4.bsnOwner = Sheet4.Range("H11")
    Set Sheet4.bsnName = Sheet4.Range("H13")
    Set Sheet4.bsnAddr = Sheet4.Range("H15")
    Set Sheet4.bsnType = Sheet4.Range("H17")
    Set Sheet4.orNumber = Sheet4.Range("H19")
    Set Sheet4.orDate = Sheet4.Range("H21")
    
    Set Sheet3.datpermitType = Sheet3.Range("A7")
    Set Sheet3.datdateIssued = Sheet3.Range("K11")
    Set Sheet3.datpermitNo = Sheet3.Range("C11")
    Set Sheet3.datbsnOwner = Sheet3.Range("F14")
    Set Sheet3.datbsnName = Sheet3.Range("F15")
    Set Sheet3.datbsnAddr = Sheet3.Range("F16")
    Set Sheet3.datbsnTypes = Sheet3.Range("D23")
    Set Sheet3.datorNumber = Sheet3.Range("J49")
    Set Sheet3.datorDate = Sheet3.Range("J50")
End Function

Public Function Protect_Toggle(protect As Boolean)
    If protect Then
        Sheet4.protect ("admin.pass")
        Sheet3.protect ("admin.pass")
        Sheet5.protect ("admin.pass")
    Else
        Sheet4.Unprotect ("admin.pass")
        Sheet3.Unprotect ("admin.pass")
        Sheet5.Unprotect ("admin.pass")
    End If
End Function

Public Function CopyData()
    Dim businesstypes() As String
    Dim i As Integer
    Dim totalpayment As Double

    PrepareDatas
    
    businesstypes = Split(Sheet4.bsnType, ",")
    
    Sheet3.datpermitType.Value = Sheet4.permitType.Value
    Sheet3.datdateIssued.Value = Sheet4.dateIssued.Value
    Sheet3.datpermitNo.Value = Sheet4.permitNo.Value
    Sheet3.datbsnOwner.Value = UCase(Sheet4.bsnOwner.Value)
    Sheet3.datbsnName.Value = UCase(Sheet4.bsnName.Value)
    Sheet3.datbsnAddr.Value = UCase(Sheet4.bsnAddr.Value)
    
    Sheet3.datbsnTypes.Value = Sheet4.bsnType.Value
    
    'clear values of payments
    Sheet3.Range("A30:K40").ClearContents
    
    'cell reference increment
    i = 31
    totalpayment = 0

On Error GoTo msgerror

    'include now payment descriptions
    If frmAddMore.pval1 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "PERMIT FEE"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval1, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval1
    End If
    
    If frmAddMore.pval2 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "LICENSE FEE"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval2, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval2
    End If
    
    If frmAddMore.pval3 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "SAN/MED"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval3, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval3
    End If
    
    If frmAddMore.pval4 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "CERT. FIRE FEE SAFETY"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval4, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval4
    End If
    
    If frmAddMore.pval5 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "WEIGHING FEE"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval5, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval5
    End If
    
    If frmAddMore.pval6 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "PENALTY"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval6, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval6
    End If
    
    If frmAddMore.pval7 <> 0 Then
        'description
        Sheet3.Range("B" & i).Value = "ZONING"
        'payment amount
        Sheet3.Range("K" & i).Value = Format(frmAddMore.pval7, "Standard")
        
        i = i + 1
        totalpayment = totalpayment + frmAddMore.pval7
    End If
    
    'show total
    Sheet3.Range("I" & i + 1).Value = "TOTAL"
    Sheet3.Range("K" & i + 1).Value = totalpayment
    
    
    Sheet3.datorNumber.Value = Sheet4.orNumber.Value
    Sheet3.datorDate.Value = Format(Sheet4.orDate.Value, "mm/dd/yyyy")
    
msgerror:
    If Err.Number > 0 Then
        MsgBox "Please do not leave every field empty." & vbNewLine & "If payment is unable, input 0." & vbNewLine & vbNewLine & Err.Description, vbCritical
    End If
    
End Function

Public Function RecordData()
    PrepareDatas
    
    If Sheet4.permitNo.Value = Sheet5.Range("C5").Value Then
    
        If MsgBox("The same PERMIT NUMBER is recently recorded." & vbNewLine & "Do you want overwrite current record?" & vbNewLine & "Current will be deleted.", vbYesNoCancel) = vbYes Then
            'permit type
            Sheet5.Range("A5").Value = Sheet4.permitType
            'date issued
            Sheet5.Range("B5").Value = Format(Sheet4.dateIssued, "mm/dd/yyyy")
            'permit number
            Sheet5.Range("C5").Value = Sheet4.permitNo
            'owner
            Sheet5.Range("D5").Value = UCase(Sheet4.bsnOwner)
            'name
            Sheet5.Range("E5").Value = UCase(Sheet4.bsnName)
            'address
            Sheet5.Range("F5").Value = UCase(Sheet4.bsnAddr)
            'type
            Sheet5.Range("G5").Value = Sheet4.bsnType
            'o.r. number
            Sheet5.Range("H5").Value = Sheet4.orNumber
            'o.r. date
            Sheet5.Range("I5").Value = Format(Sheet4.orDate, "mm/dd/yyyy")
            'no. of employees
            Sheet5.Range("J5").Value = frmAdditionalInfo.empqtty
            'capital investment
            Sheet5.Range("K5").Value = frmAdditionalInfo.capitalInvest
            
            MsgBox "Business Permit is successfully recorded!", vbInformation
        End If
    Else
        Sheet5.Range("A5").EntireRow.Insert
        
        'permit type
        Sheet5.Range("A5").Value = Sheet4.permitType
        'date issued
        Sheet5.Range("B5").Value = Format(Sheet4.dateIssued, "mm/dd/yyyy")
        'permit number
        Sheet5.Range("C5").Value = Sheet4.permitNo
        'owner
        Sheet5.Range("D5").Value = UCase(Sheet4.bsnOwner)
        'name
        Sheet5.Range("E5").Value = UCase(Sheet4.bsnName)
        'address
        Sheet5.Range("F5").Value = UCase(Sheet4.bsnAddr)
        'type
        Sheet5.Range("G5").Value = Sheet4.bsnType
        'o.r. number
        Sheet5.Range("H5").Value = Sheet4.orNumber
        'o.r. date
        Sheet5.Range("I5").Value = Format(Sheet4.orDate, "mm/dd/yyyy")
        'no. of employees
        Sheet5.Range("J5").Value = frmAdditionalInfo.empqtty
        'capital investment
        Sheet5.Range("K5").Value = frmAdditionalInfo.capitalInvest
    
        MsgBox "Business Permit is successfully recorded!", vbInformation
    End If
End Function
