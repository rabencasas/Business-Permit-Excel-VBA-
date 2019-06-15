VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddress 
   ClientHeight    =   4410
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8055
   OleObjectBlob   =   "frmAddress.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    If ListBox1.ListIndex = 9 Then
        Sheet4.bsnAddr.Value = ListBox1.Value
    Else
        Sheet4.bsnAddr.Value = ComboBox1.Value & ", " & ListBox1.Value
    End If
    
    Sheet4.bsnType.Activate
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    ListBox1.List = Sheet4.Range("B31:B40").Value
    ComboBox1.List = Sheet4.Range("F31:F265").Value
    
    ComboBox1.SelStart = 0
    ComboBox1.SelLength = ComboBox1.TextLength
    ComboBox1.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub
