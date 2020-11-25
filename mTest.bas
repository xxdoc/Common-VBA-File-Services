Attribute VB_Name = "mTest"
Option Explicit
' ------------------------------------------------------------
' Standard Module mTest Test of all Existence checks variants
'                       in module mExists
' -----------------------------------------------------------
'Declare API
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
Const CONCAT = "||"
Const SHIFT_KEY = 16

Private Sub Regression_Test()
' ---------------------------------------------------------
' All results are asserted and there is no intervention
' required for the whole test. When an assertion fails the
' test procedure will stop and indicates the problem with
' the called procedure.
' An execution trace is displayed for each test procedure.
' ---------------------------------------------------------
    Const PROC = "Test_All"

    On Error GoTo eh
    
    Test_FileExists

xt: Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub

Public Sub Test_SelectFile()
Dim fso As File

    If mFile.SelectFile("", "*.xl*", "Excel File", fso) = True Then
        Debug.Assert fso.Path = ThisWorkbook.FullName
    Else
        Debug.Assert fso Is Nothing
    End If
    
End Sub

Private Sub Test_FileExists()
Const PROC      As String = "Test_FileExists"       ' This procedure's name for the error handling and execution tracking
Dim wb          As Workbook
Dim fso         As File
Dim fsoExists   As File
Dim cltOfFiles  As Collection

    On Error GoTo eh
    BoP ErrSrc(PROC)
    Set wb = ThisWorkbook
    
    With New FileSystemObject
        Set fso = .GetFile(wb.FullName)
    End With
    
    '~~ 1. File by object (exists)
    Debug.Assert mFile.Exists(fso) = True
'    Debug.Assert fsoExists Is fso
    
    '~~ 2. File by Fullname (exists)
    Debug.Assert mFile.Exists(fso.Path, fsoExists) = True
    Debug.Assert fsoExists Is fso
    
    '~~ 3. File by Fullname with wildcard * (exactly one exists)
    Debug.Assert mFile.Exists(left(fso.Path, Len(fso.Path) - 1) & "*", , cltOfFiles) = True
    Debug.Assert cltOfFiles.Count = 1
    Debug.Assert cltOfFiles.Item(1).Path = fso.Path
    
    '~~ 4. File by Fullname with wildcard * (2 such files exist)
    Debug.Assert mFile.Exists(wb.Path & "\fMsg*", , cltOfFiles) = True
    Debug.Assert cltOfFiles.Count = 2
    Debug.Assert cltOfFiles.Item(1).Name = "fMsg.frm"
    Debug.Assert cltOfFiles.Item(2).Name = "fMsg.frx"
    
    '~~ 5. File by Fullname with wildcard * (2 such files exist but in a sub-folder)
    Debug.Assert mFile.Exists(Replace(wb.Path & "\fMsg*", "\" & Split(wb.Name, ".")(0), vbNullString), , cltOfFiles) = True
    Debug.Assert cltOfFiles.Count >= 2
    Debug.Assert cltOfFiles.Item(1).Name = "fMsg.frm"
    Debug.Assert cltOfFiles.Item(2).Name = "fMsg.frx"
    
    '~~ 6. File not exists
    Debug.Assert mFile.Exists("Test.txt") = False

    '~~ 7. Neither a File object nor a string
    On Error Resume Next
    mFile.Exists wb
    Debug.Assert AppErr(err.Number) = 1
    On Error GoTo eh
        
xt:
    EoP ErrSrc(PROC)
    Exit Sub
    
eh:
    ErrMsg ErrSrc(PROC)
End Sub
  
Public Sub Test_ToArray()
Const PROC = "Test_ToArray"
Dim sFile As String
Dim fl      As File
Dim a       As Variant
Dim v       As Variant

    On Error GoTo eh
    BoP ErrSrc(PROC)
    
    sFile = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mBasic.bas"
    With New FileSystemObject
        Set fl = .GetFile(sFile)
    End With
    a = mFile.ToArray(fl)
    
#If Debugging Then
    For Each v In a
        Debug.Print ">>" & v & "<<"
    Next v
#End If

xt:
    EoP ErrSrc(PROC)
    Exit Sub
    
eh: ErrMsg ErrSrc(PROC)
End Sub


Private Sub ErrMsg( _
             ByVal err_source As String, _
    Optional ByVal err_no As Long = 0, _
    Optional ByVal err_dscrptn As String = vbNullString, _
    Optional ByVal err_line As Long = 0, _
    Optional ByVal err_asserted = 0)
' --------------------------------------------------
' Note! Because the mTrc trace module is an optional
'       module of the mErH error handler module it
'       cannot use the mErH's ErrMsg procedure and
'       thus uses its own - with the known
'       disadvantage that the title maybe truncated.
' --------------------------------------------------
    Dim sTitle      As String
    Dim sDetails    As String
    
    If err_no = 0 Then err_no = err.Number
    If err_dscrptn = vbNullString Then err_dscrptn = err.Description
    If err_line = 0 Then err_line = Erl
    
    ErrMsgMatter err_source:=err_source, err_no:=err_no, err_line:=err_line, err_dscrptn:=err_dscrptn, msg_title:=sTitle, msg_details:=sDetails
    
#If Test Then
    If err_no <> err_asserted _
    Then MsgBox Prompt:="Error description:" & vbLf & _
                        err_dscrptn & vbLf & vbLf & _
                        "Error source/details:" & vbLf & _
                        sDetails, _
                buttons:=vbOKOnly, _
                Title:=sTitle
#Else
    MsgBox Prompt:="Error description:" & vbLf & _
                    err_dscrptn & vbLf & vbLf & _
                   "Error source/details:" & vbLf & _
                   sDetails, _
           buttons:=vbOKOnly, _
           Title:=sTitle
#End If
End Sub

Private Sub ErrMsgMatter(ByVal err_source As String, _
                         ByVal err_no As Long, _
                         ByVal err_line As Long, _
                         ByVal err_dscrptn As String, _
                Optional ByRef msg_title As String, _
                Optional ByRef msg_type As String, _
                Optional ByRef msg_line As String, _
                Optional ByRef msg_no As Long, _
                Optional ByRef msg_details As String, _
                Optional ByRef msg_dscrptn As String, _
                Optional ByRef msg_info As String)
' ---------------------------------------------------------------------------------
' Returns all matter to build a proper error message.
' msg_line:    at line <err_line>
' msg_no:      1 to n (an Application error translated back into its origin number)
' msg_title:   <error type> <error number> in:  <error source>
' msg_details: <error type> <error number> in <error source> [(at line <err_line>)]
' msg_dscrptn: the error description
' msg_info:    any text which follows the description concatenated by a ||
' ---------------------------------------------------------------------------------
    If InStr(1, err_source, "DAO") <> 0 _
    Or InStr(1, err_source, "ODBC Teradata Driver") <> 0 _
    Or InStr(1, err_source, "ODBC") <> 0 _
    Or InStr(1, err_source, "Oracle") <> 0 Then
        msg_type = "Database Error "
    Else
      msg_type = IIf(err_no > 0, "VB-Runtime Error ", "Application Error ")
    End If
   
    msg_line = IIf(err_line <> 0, "at line " & err_line, vbNullString)
    msg_no = IIf(err_no < 0, err_no - vbObjectError, err_no)
    msg_title = msg_type & msg_no & " in:  " & err_source
    msg_details = IIf(err_line <> 0, msg_type & msg_no & " in: " & err_source & " (" & msg_line & ")", msg_type & msg_no & " in " & err_source)
    msg_dscrptn = IIf(InStr(err_dscrptn, CONCAT) <> 0, Split(err_dscrptn, CONCAT)(0), err_dscrptn)
    If InStr(err_dscrptn, CONCAT) <> 0 Then msg_info = Split(err_dscrptn, CONCAT)(1) Else msg_info = vbNullString

End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mTest" & ">" & sProc
End Function
