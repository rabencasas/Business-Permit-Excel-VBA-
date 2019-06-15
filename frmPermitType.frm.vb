VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPermitType 
   ClientHeight    =   1545
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   6675
   OleObjectBlob   =   "frmPermitType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPermitType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdNew_Click()
    Protect_Toggle (False)
    
    PrepareDatas
    
    Sheet4.permitType.Value = "NEW"
    Hide
    
    Protect_Toggle (True)
End Sub

Private Sub cmdRenewal_Click()
    Protect_Toggle (False)
    
    PrepareDatas
    
    Sheet4.permitType.Value = "RENEWAL"
    Hide
    
    Protect_Toggle (True)
End Sub
