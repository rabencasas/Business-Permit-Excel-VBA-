VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmPrintCopy 
   ClientHeight    =   2490
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   5580
   OleObjectBlob   =   "frmPrintCopy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmPrintCopy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Private Sub CommandButton1_Click()
    
    CopyData
    
    Sheet3.ExportAsFixedFormat Type:=xlTypePDF, Filename:=ActiveWorkbook.Path & "/businesspermit.pdf"
    
    If CheckBox1.Value Then
        ActiveWorkbook.FollowHyperlink ActiveWorkbook.Path & "/businesspermit.pdf"
        'Call Shell(ActiveWorkbook.Path & "/businesspermit.pdf", vbNormalFocus)
    Else
        Sheet3.PrintOut from:=1, To:=1, copies:=TextBox1.Value
    End If
    
    Hide
    
End Sub

Private Sub UserForm_Activate()
    TextBox1.SetFocus

End Sub

Private Sub UserForm_Click()

End Sub
