VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmSetDate 
   Caption         =   "Set Date Issued (mm/dd/yyyy):"
   ClientHeight    =   1335
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6675
   OleObjectBlob   =   "frmSetDate.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmSetDate"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSetToday_Click()
    TextBox1.Value = Format(Now, "mm/dd/yyyy")
End Sub

Private Sub CommandButton1_Click()
    Sheet4.Range("H8").Value = TextBox1.Value
    Sheet4.Range("H10").Activate
    Me.Hide
End Sub

Private Sub UserForm_Activate()
    TextBox1.Value = Sheet4.Range("H8").Value

    TextBox1.SelStart = 0
    TextBox1.SelLength = TextBox1.TextLength
    TextBox1.SetFocus
End Sub

Private Sub UserForm_Click()

End Sub
