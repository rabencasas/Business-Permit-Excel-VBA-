VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmObligationNoDialog 
   Caption         =   "Set obligation number format:"
   ClientHeight    =   1635
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8955
   OleObjectBlob   =   "frmObligationNoDialog.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmObligationNoDialog"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton1_Click()
    Sheet4.Range("H10").Value = txtPart1.Value & "-" & txtPart2.Value & "-" & txtPart3.Value & "-" & txtPart4.Value
    Me.Hide
End Sub
