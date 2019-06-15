VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet4"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Public permitType, dateIssued, permitNo, bsnOwner, bsnName, bsnAddr, bsnType, orNumber, orDate As Range

Private Sub cmdAdd_Click()
    Protect_Toggle (False)
    
    frmAddMore.Show
    
    Protect_Toggle (True)
End Sub

Private Sub cmdPreview_Click()
    Protect_Toggle (False)
    
    CopyData
    
    Sheet3.PrintPreview
    
    Protect_Toggle (True)
End Sub

Private Sub cmdPrint_Click()
    Protect_Toggle (False)

    frmPrintCopy.Show
    
    Protect_Toggle (True)
End Sub

Private Sub cmdRecord_Click()
    Protect_Toggle (False)

    RecordData
    
    Protect_Toggle (True)
End Sub

Private Sub CommandButton1_Click()
    
    frmAddress.Show
    
End Sub

Private Sub CommandButton2_Click()

    frmPermitType.Show
    
End Sub

Private Sub CommandButton3_Click()
    PrepareDatas
    
    If permitNo.Value <> "" Then
        permitNo.Value = permitNo.Value + 1
    End If
End Sub

Private Sub CommandButton4_Click()
    frmAdditionalInfo.Show
End Sub

Private Sub CommandButton5_Click()
    PrepareDatas
    
    dateIssued.Value = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub CommandButton7_Click()
    PrepareDatas
    
    If permitNo.Value <> "" Then
        permitNo.Value = permitNo.Value - 1
    End If
End Sub
