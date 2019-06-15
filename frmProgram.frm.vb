VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmProgram 
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9660
   OleObjectBlob   =   "frmProgram.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmProgram"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    OpenSheet

    If ListBox1.Value <> "" Then
        Sheet4.Range("H20").Value = ListBox1.ListIndex + 1 & "_" & ListBox1.Value
        Me.Hide
    Else
        MsgBox "Please select a valid program.", vbOKOnly, "Invalid Selection"
        ListBox1.ListIndex = 0
    End If
    
    CloseSheet
End Sub

Private Sub ListBox1_Change()
    
    lblBalance.Caption = "BALANCE (PHP): " & Format(Sheet3.Range("N" & ListBox1.ListIndex + 9).Value, "Standard")

End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    OpenSheet
    
    If ListBox1.Value <> "" Then
        Sheet4.Range("H20").Value = ListBox1.ListIndex + 1 & "_" & ListBox1.Value
        Me.Hide
    Else
        MsgBox "Please select a valid program.", vbOKOnly, "Invalid Selection"
        ListBox1.ListIndex = 0
    End If
    
    CloseSheet
End Sub

Private Sub UserForm_Activate()
    ListBox1.List = Sheet3.Range("D9:D70").Value
End Sub

Private Sub UserForm_Click()
    
End Sub
