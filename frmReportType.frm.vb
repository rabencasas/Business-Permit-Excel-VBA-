VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmReportType 
   ClientHeight    =   6540
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   10245
   OleObjectBlob   =   "frmReportType.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmReportType"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub CommandButton5_Click()
    Me.Hide
    Sheet3.PrintPreview
    Me.Show
End Sub

Private Sub ListBox1_Click()

End Sub

Private Sub cmdGenerateReport_Click()
    If lstReportType.ListIndex <> -1 Then
        GenerateReport (lstReportType.ListIndex)
    Else
        MsgBox "Please select the report types.", vbCritical, "No report type selected"
    End If
End Sub

Private Sub cmdPreview_Click()
    Hide
    Sheet1.PrintPreview
    Me.Show
End Sub

Private Sub UserForm_Activate()
    lstReportType.List = Sheet3.Range("D80:D83").Value
End Sub

Private Sub UserForm_Click()
    
End Sub
