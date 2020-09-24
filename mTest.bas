Attribute VB_Name = "mTest"
Option Explicit
' ------------------------------------------------------------
' Standard Module mTest Test of all Existence checks variants
'                       in module mExists
' -----------------------------------------------------------
'Declare API
Declare PtrSafe Function GetKeyState Lib "user32" (ByVal vKey As Integer) As Integer
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

    On Error GoTo on_error
    
    Test_FileExists

exit_proc:
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
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

    On Error GoTo on_error
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
    Debug.Assert mFile.Exists(Left(fso.Path, Len(fso.Path) - 1) & "*", , cltOfFiles) = True
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
    Debug.Assert AppErr(Err.Number) = 1
    On Error GoTo on_error
        
exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub
  
Public Sub Test_ToArray()
Const PROC = "Test_ToArray"
Dim sFile As String
Dim fl      As File
Dim a       As Variant
Dim v       As Variant

    On Error GoTo on_error
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

exit_proc:
    EoP ErrSrc(PROC)
    Exit Sub
    
on_error:
    mErrHndlr.ErrHndlr Err.Number, ErrSrc(PROC), Err.Description, Erl
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.Name & ">mTest" & ">" & sProc
End Function
