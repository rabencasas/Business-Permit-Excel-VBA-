VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAdditionalInfo 
   Caption         =   "UserForm1"
   ClientHeight    =   4290
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8130
   OleObjectBlob   =   "frmAdditionalInfo.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAdditionalInfo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public empqtty, capitalInvest As Double

Private Sub cmdAdd_Click()
    If txtemployeesqtty.Value <> 0 Or txtemployeesqtty.Value <> "" Then
        empqtty = txtemployeesqtty.Value
    Else
        empqtty = 0
    End If
    
    If txtcapitalinvestment.Value <> 0 Or txtcapitalinvestment.Value <> "" Then
        capitalInvest = txtcapitalinvestment.Value
    Else
        capitalInvest = 0
    End If
    
    Hide
End Sub
