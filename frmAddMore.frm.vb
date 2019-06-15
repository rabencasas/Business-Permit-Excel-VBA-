VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddMore 
   Caption         =   "Add Payment Descriptions"
   ClientHeight    =   9015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   6705
   OleObjectBlob   =   "frmAddMore.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddMore"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public pval1, pval2, pval3, pval4, pval5, pval6, pval7 As Double

Private Sub cmdAdd_Click()
    If payment1.Value <> 0 Or payment1.Value <> "" Then
        pval1 = payment1.Value
    Else
        pval1 = 0
    End If
    
    If payment2.Value <> 0 Or payment2.Value <> "" Then
        pval2 = payment2.Value
    Else
        pval2 = 0
    End If
    
    If payment3.Value <> 0 Or payment3.Value <> "" Then
        pval3 = payment3.Value
    Else
        pval3 = 0
    End If
    
    If payment4.Value <> 0 Or payment4.Value <> "" Then
        pval4 = payment4.Value
    Else
        pval4 = 0
    End If
    
    If payment5.Value <> 0 Or payment5.Value <> "" Then
        pval5 = payment5.Value
    Else
        pval5 = 0
    End If
    
    If payment6.Value <> 0 Or payment6.Value <> "" Then
        pval6 = payment6.Value
    Else
        pval6 = 0
    End If
    
    If payment7.Value <> 0 Or payment7.Value <> "" Then
        pval7 = payment7.Value
    Else
        pval7 = 0
    End If
    
    Hide
End Sub
