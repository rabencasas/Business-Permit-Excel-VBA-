VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSearchDialog 
   ClientHeight    =   3195
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   9285
   OleObjectBlob   =   "frmSearchDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSearchDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCancelUpdate1_Click()
    If cmdCancelUpdate1.Caption = "CANCEL OBLIGATION" Then
        If MsgBox("Are you sure you want to cancel the obligation " & txtObligationNo.Value & "?", vbYesNoCancel, "Confirm cancel obligation") = vbYes Then
        
        End If
    Else
        
    End If
End Sub

Private Sub cmdSearch_Click()
    
    If txtSearchValue.Value <> "" Then
        Dim result As Range
        
        If optObNo.Value = True Then
            Set result = Sheet5.Range("C:C").Find(What:=txtSearchValue.Value, After:=Sheet5.Range("C4"), SearchDirection:=xlNext)
        ElseIf optVoucherName.Value = True Then
            Set result = Sheet5.Range("D:D").Find(txtSearchValue.Value, Sheet5.Range("D4"))
        Else
            Set result = Sheet5.Range("I:I").Find(txtSearchValue.Value, Sheet5.Range("I4"))
        End If
        
        If result Is Nothing Then
            MsgBox "Obligation was not found.", vbCritical, "Obligation not found"
        Else
            Sheet4.Range("H8").Value = Format(Sheet5.Range("A" & result.Row).Value, "mm/dd/yyyy")
            Sheet4.Range("H10").Value = Sheet5.Range("C" & result.Row).Value
            Sheet4.Range("H12").Value = Sheet5.Range("D" & result.Row).Value
            Sheet4.Range("H14").Value = Sheet5.Range("P" & result.Row).Value
            Sheet4.Range("H16").Value = Sheet5.Range("I" & result.Row).Value
            Sheet4.Range("H18").Value = Format(Sheet5.Range("N" & result.Row).Value, "Standard")
            Sheet4.Range("H20").Value = Sheet5.Range("T" & result.Row).Value
            
            MsgBox "One result found", vbInformation, "One result found"
            
            Sheet4.cmdCloseSearch.Visible = True
            Sheet4.cmdPreview.Caption = "OPTIONS"
            Sheet4.Range("H1").Value = "*** SEARCH MODE ***"
            
            Sheet4.Range("E24").Value = ""
            Sheet4.Range("E25").Value = ""
            
            Sheet4.resultRowNo = result.Row
            
            'Me.Hide
        End If
    Else
        MsgBox "Please input the obligation number first.", vbCritical, "Invalid obligatoin number"
    End If
    
End Sub

Private Sub cmdUpdate2_Click()
    If cmdUpdate2.Caption = "UPDATE OBLIGATION" Then
        If MsgBox("Are you sure you want to update current obligation?", vbYesNoCancel, "Confirm update obligation") = vbYes Then
            txtStatus.Enabled = True
            txtName.Enabled = True
            txtAddress.Enabled = True
            txtBeneficiary.Enabled = True
            txtAmount.Enabled = True
            txtCategory.Enabled = True
            
            cmdCancelUpdate1.Caption = "SAVE CHANGES"
            cmdUpdate2.Caption = "CANCEL"
        End If
    Else
        
    End If
End Sub

Private Sub CommandButton1_Click()
    frmSearchType.Show
End Sub

Private Sub Label3_Click()

End Sub

Private Sub optObNo_Click()

End Sub

Private Sub UserForm_Activate()
    optObNo.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub
