VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetClerk 
   Caption         =   "Set clerk name:"
   ClientHeight    =   1365
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   5580
   OleObjectBlob   =   "frmSetClerk.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetClerk"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Sheet4.Range("H5").Value = TextBox1.Value
    Sheet4.Range("H10").Activate
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    TextBox1.Value = Sheet4.Range("H5").Value

    TextBox1.SelStart = 0
    TextBox1.SelLength = TextBox1.TextLength
    TextBox1.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub
