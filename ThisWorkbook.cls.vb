VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "ThisWorkbook"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Workbook_Open()
    
    Protect_Toggle (False)
    
    PrepareDatas
    
    'automatically set date
    Sheet4.dateIssued.Value = Format(DateTime.Now, "mm/dd/yyyy")
    
    'automatically set permit number
    Sheet4.permitNo.Value = Sheet5.Range("C5").Value + 1
        
    'clear all other datas
    Sheet4.permitType.Value = ""
    Sheet4.bsnOwner.Value = ""
    Sheet4.bsnName.Value = ""
    Sheet4.bsnAddr.Value = ""
    Sheet4.bsnType.Value = ""
    Sheet4.orNumber.Value = ""
    Sheet4.orDate.Value = ""
    
    Sheet4.Activate
    Sheet4.permitType.Activate
    
    Protect_Toggle (True)
End Sub
