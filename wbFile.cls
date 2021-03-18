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
                                    ' ----------------------------------------
Private Const HOSTED_RAWS = "mFile" ' The Common Component hosted, developed,
                                    ' tested, and provided by this Workbook
                                    ' ----------------------------------------

Private Sub Workbook_Open()
    
    '~~ ------------------------------------------------------------------
    '~~ CompMan Workbook_Open service 'UpdateRawClones':
    '~~ Executed by the Addin *) or via the development instance when open
    '~~ *) automatically available only when referenced by the VB-Project
    mCompManClient.CompManService "UpdateRawClones", HOSTED_RAWS
    '~~ ------------------------------------------------------------------
    
End Sub

Private Sub Workbook_BeforeSave(ByVal SaveAsUI As Boolean, Cancel As Boolean)

    '~~ ------------------------------------------------------------------
    '~~ 'ExportChangedComponents' service, preferrably performed by the
    '~~ CompMan Addin, when not open alternatively by the open CompMan-
    '~~ Development-Instance-Workbook. When neither is open the service
    '~~ is not performed without notic.
    mCompManClient.CompManService "ExportChangedComponents", HOSTED_RAWS
    '~~ ------------------------------------------------------------------

End Sub
