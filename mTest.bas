Attribute VB_Name = "mTest"
Option Explicit
' ------------------------------------------------------------
' Standard Module mTest
' Test of all services of the module.
' -----------------------------------------------------------

Private Property Let Status(ByVal s As String)
    If s <> vbNullString Then
        Application.StatusBar = "regression test mFile: " & s
    Else
        Application.StatusBar = vbNullString
    End If
End Property

Public Sub Regression()
' ---------------------------------------------------------
' All results are asserted and there is no intervention
' required for the whole test. When an assertion fails the
' test procedure will stop and indicates the problem with
' the called procedure.
' An execution trace is displayed for each test procedure.
' ---------------------------------------------------------
    Const PROC = "Regression"

    On Error GoTo eh
    Dim sTestStatus As String
    
    sTestStatus = "mFile Regression Test: "

    mErH.BoTP ErrSrc(PROC), mErH.AppErr(1) ' For the very last test on an error condition
    mTest.Test_01_FileExists_Not
    mTest.Test_02_FileExists_ByObject
    mTest.Test_03_FileExists_ByFullName
    mTest.Test_04_FileExists_ByFullName_WildCard_ExactlyOne
    mTest.Test_05_FileExists_ByFullName_WildCard_MoreThanOne
    mTest.Test_06_FileExists_WildCard_MoreThanOne_InSubFolder
    mTest.Test_07_SelectFile
    mTest.Test_08_Arry_Get
    mTest.Test_09_File_Differs
    mTest.Test_10_Txt
    Test_99_FileExists_NoFileObject_NoString ' Error AppErr(1) !
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_07_SelectFile()
    Const PROC = "Test_07_SelectFile"
    
    On Error GoTo eh
    Dim fso As File

    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    If mFile.SelectFile( _
                        sel_init_path:=ThisWorkbook.Path, _
                        sel_filters:="*.xl*", _
                        sel_filter_name:="Excel File", _
                        sel_title:="Select the (preselected by filtering) file", _
                        sel_result:=fso _
                        ) = True Then
        Debug.Assert fso.Path = ThisWorkbook.FullName
    Else
        Debug.Assert fso Is Nothing
    End If
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_02_FileExists_ByObject()
    Const PROC      As String = "Test_02_FileExists_ByObject"
    
    On Error GoTo eh
    Dim wb          As Workbook
    Dim fso         As File

    Status = ErrSrc(PROC)
    Set wb = ThisWorkbook
    With New FileSystemObject
        Set fso = .GetFile(wb.FullName)
    End With
    
    mErH.BoP ErrSrc(PROC), "xst_file:=", wb.FullName
    Debug.Assert mFile.Exists(xst_file:=fso) = True
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub
  

Public Sub Test_03_FileExists_ByFullName()
    Const PROC      As String = "Test_03_FileExists_ByFullName"
    
    On Error GoTo eh
    Dim wb          As Workbook
    Dim fso         As File
    Dim fsoExists   As File

    mErH.BoP ErrSrc(PROC)
    Set wb = ThisWorkbook
    Status = ErrSrc(PROC)
    
    With New FileSystemObject
        Set fso = .GetFile(wb.FullName)
    End With
      
    Debug.Assert mFile.Exists(fso.Path, fsoExists) = True
    Debug.Assert fsoExists Is fso
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_99_FileExists_NoFileObject_NoString()
    Const PROC = "Test_99_FileExists_NoFileObject_NoString"
    
    On Error GoTo eh

    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    mFile.Exists ThisWorkbook
    Debug.Assert mErH.AppErr(1)
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_01_FileExists_Not()
    Const PROC = "Test_01_FileExists_Not"
    Const NOT_EXIST = "Test.txt"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC), "xst_file:=", NOT_EXIST
    Status = ErrSrc(PROC)
    Debug.Assert mFile.Exists(xst_file:="Test.txt") = False
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_04_FileExists_ByFullName_WildCard_ExactlyOne()
    Const PROC = "Test_04_FileExists_ByFullName_WildCard_ExactlyOne"
    
    On Error GoTo eh
    Dim wb      As Workbook
    Dim fsoFile As File
    Dim fso     As New FileSystemObject
    Dim cll     As Collection
    Dim sWldCrd As String
    
    Status = ErrSrc(PROC)
    ' Prepare
    Set wb = ThisWorkbook
    Set fsoFile = fso.GetFile(wb.FullName)
    sWldCrd = VBA.Left$(fsoFile.Path, Len(fsoFile.Path) - 3) & "*"
    
    ' Test
    mErH.BoP ErrSrc(PROC), "xst_file:=", sWldCrd
    Debug.Assert mFile.Exists(xst_file:=sWldCrd, xst_cll:=cll) = True
    Debug.Assert cll.Count = 1
    Debug.Assert cll.Item(1).Path = fsoFile.Path
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_06_FileExists_WildCard_MoreThanOne_InSubFolder()
    Const PROC      As String = "Test_06_FileExists_WildCard_MoreThanOne_InSubFolder"       ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim wb          As Workbook
    Dim cllFiles    As Collection
    Dim sWldCrd     As String

    Status = ErrSrc(PROC)
    ' Prepare
    Set wb = ThisWorkbook
    sWldCrd = Replace(wb.Path & "\fMsg*", "\" & Split(wb.name, ".")(0), vbNullString)
    
    ' Test
    mErH.BoP ErrSrc(PROC), "xst_file:=", sWldCrd
    Debug.Assert mFile.Exists( _
                              xst_file:=sWldCrd, _
                              xst_cll:=cllFiles _
                             ) = True
    Debug.Assert cllFiles.Count >= 2
    Debug.Assert cllFiles.Item(1).name = "fMsg.frm"
    Debug.Assert cllFiles.Item(2).name = "fMsg.frx"
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_05_FileExists_ByFullName_WildCard_MoreThanOne()
    Const PROC = "Test_05_FileExists_ByFullName_WildCard_MoreThanOne"
    
    On Error GoTo eh
    Dim wb          As Workbook
    Dim cllFiles    As Collection
    Dim sWldCrd     As String

    Status = ErrSrc(PROC)
    ' Prepare
    Set wb = ThisWorkbook
    sWldCrd = wb.Path & "\fMsg*"
    
    ' Test
    mErH.BoP ErrSrc(PROC), "xst_file:=", sWldCrd
    Debug.Assert mFile.Exists(xst_file:=sWldCrd, xst_cll:=cllFiles) = True
    Debug.Assert cllFiles.Count = 2
    Debug.Assert cllFiles.Item(1).name = "fMsg.frm"
    Debug.Assert cllFiles.Item(2).name = "fMsg.frx"
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub
  
Public Sub Test_08_Arry_Get()
    Const PROC = "Test_08_Arry_Get"
    
    On Error GoTo eh
    Dim sFile       As String
    Dim lInclEmpty  As Long
    Dim lEmpty1     As Long
    Dim lExclEmpty  As Long
    Dim lEmpty2     As Long
    Dim fso         As New FileSystemObject
    Dim a           As Variant
    Dim v           As Variant
    Dim sTemp       As String
    
    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    sFile = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas"
    
    sTemp = mFile.Temp()
    mFile.Txt(sTemp) = "xxx" & vbCrLf & "" & "yyy"
    
    '~~ Count empty records when array contains all text lines
    a = mFile.Arry(fa_file_full_name:=sTemp, fa_exclude_empty_records:=False)
    lInclEmpty = UBound(a) + 1
    lEmpty1 = 0
    For Each v In a
        If VBA.Trim$(v) = vbNullString Then lEmpty1 = lEmpty1 + 1
        If VBA.Len(Trim$(v)) = 0 Then lEmpty2 = lEmpty2 + 1
    Next v
    
    '~~ Count empty records
    a = mFile.Arry(fa_file_full_name:=sTemp, fa_exclude_empty_records:=True)
    lExclEmpty = UBound(a) + 1
    Debug.Assert lExclEmpty = lInclEmpty - lEmpty1
    
xt: If fso.FileExists(sTemp) Then fso.DeleteFile sTemp
    Set fso = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub
  
Public Sub Test_09_File_Differs_False()
    Const PROC = "Test_09_File_Differs"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim f1      As File
    Dim f2      As File
    Dim i       As Long
    Dim dctDiff As Dictionary
    
    Status = ErrSrc(PROC)
    ' Prepare
    sFile = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas"
    Set f1 = fso.GetFile("E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas")
    Set f2 = fso.GetFile("E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas")
    
    ' Test
    mErH.BoP ErrSrc(PROC), "dif_file1 = ", f1.name, "dif_file2 = ", f2.name
    Set dctDiff = mFile.Differs(dif_file1:=f1, dif_file2:=f2, dif_ignore_empty_records:=True)
    Debug.Assert dctDiff.Count = 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_09_File_Differs()
    Const PROC = "Test_09_File_Differs"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim f1      As File
    Dim f2      As File
    Dim i       As Long
    Dim dctDiff As Dictionary
    Dim v       As Variant
    Dim sF1     As String
    Dim sF2     As String

    Status = ErrSrc(PROC)
    
    sF1 = mFile.Temp
    sF2 = mFile.Temp

    ' Prepare
    mFile.Txt(tx_file_full_name:=sF1, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFile.Txt(tx_file_full_name:=sF2, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)

    ' Test 1: Differs.Count = 0
    mErH.BoP ErrSrc(PROC)
    Set dctDiff = mFile.Differs(dif_file1:=f1 _
                              , dif_file2:=f2 _
                              , dif_ignore_empty_records:=True _
                              , dif_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 0

    ' Test 2: Differs.Count = 1
    mFile.Txt(tx_file_full_name:=sF1, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFile.Txt(tx_file_full_name:=sF2, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFile.Differs(dif_file1:=f1 _
                              , dif_file2:=f2 _
                              , dif_ignore_empty_records:=True _
                              , dif_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    For i = 0 To dctDiff.Count - 1
        Debug.Print dctDiff.Items()(i)
    Next i
    
    ' Test 3: Differs.Count = 1
    mFile.Txt(tx_file_full_name:=sF1, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    mFile.Txt(tx_file_full_name:=sF2, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFile.Differs(dif_file1:=f1 _
                              , dif_file2:=f2 _
                              , dif_ignore_empty_records:=True _
                              , dif_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    For i = 0 To dctDiff.Count - 1
        Debug.Print dctDiff.Items()(i)
    Next i
    
    ' Test 4: Differs.Count = 1
    mFile.Txt(tx_file_full_name:=sF1, tx_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFile.Txt(tx_file_full_name:=sF2, tx_append:=False) = "A" & vbCrLf & "X" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFile.Differs(dif_file1:=f1 _
                              , dif_file2:=f2 _
                              , dif_ignore_empty_records:=True _
                              , dif_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    For i = 0 To dctDiff.Count - 1
        Debug.Print dctDiff.Items()(i)
    Next i

xt: mErH.EoP ErrSrc(PROC)
    If fso.FileExists(sF1) Then fso.DeleteFile (sF1)
    If fso.FileExists(sF2) Then fso.DeleteFile (sF2)
    Exit Sub
    
eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_10_Txt()
    Const PROC = "Test_10_Txt"
    
    On Error GoTo eh
    Dim sFl     As String
    Dim sTest   As String
    Dim sResult As String
    Dim a()     As String
    Dim sSplit  As String
    Dim fso     As New FileSystemObject
    
    Status = ErrSrc(PROC)
    sFl = mFile.Temp()
    sTest = "My string"

    mFile.Txt(tx_file_full_name:=sFl _
            , tx_append:=False _
             ) = sTest
    sResult = mFile.Txt(tx_file_full_name:=sFl, tx_split:=sSplit)
    a = Split(sResult, sSplit)
    Debug.Assert a(0) = sTest
    If fso.FileExists(sFl) Then fso.DeleteFile (sFl)
    
    sFl = mFile.Temp()
    sTest = ""
    mFile.Txt(tx_file_full_name:=sFl _
            , tx_append:=False _
             ) = ""
    sResult = mFile.Txt(tx_file_full_name:=sFl, tx_split:=sSplit)
    Debug.Assert sResult = vbNullString

xt: If fso.FileExists(sFl) Then fso.DeleteFile (sFl)
    Set fso = Nothing
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Public Sub Test_11_Search()
    Const PROC = "Test_11_Search"
    
    On Error GoTo eh
    Dim cll As Collection
    Dim v   As Variant
    
    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    Set cll = mFile.Search(fs_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                         , fs_mask:="*CompMan*.frx" _
                         , fs_stop_after:=5 _
                          )
    For Each v In cll
        Debug.Print v
    Next v

    Set cll = mFile.Search(fs_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                         , fs_mask:="*CompMan*.frx" _
                         , fs_stop_after:=5 _
                         , fs_in_subfolders:=False _
                          )
    Debug.Assert cll.Count = 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case mErH.ErrMsg(ErrSrc(PROC))
        Case mErH.DebugOpt1ResumeError: Stop: Resume
        Case mErH.DebugOpt2ResumeNext: Resume Next
    End Select
End Sub

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.name & ": mTest." & sProc
End Function
