VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "wbFile"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Option Explicit

Private Const HOSTED_RAWS = "mFile"

Private Sub Workbook_Open()
#If CompMan Then
    mCompMan.UpdateClonesTheRawHasChanged uc_wb:=ThisWorkbook, uc_hosted:=HOSTED_RAWS
#End If
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)
#If CompMan Then
    mCompMan.ExportChangedComponents ec_wb:=ThisWorkbook, ec_hosted:=HOSTED_RAWS
#End If
End Sub
