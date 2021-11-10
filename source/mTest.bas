Attribute VB_Name = "mTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mTest: Test of all services of the module.
'
' Attention! To run a complete Regression_PrivateProfile test
'            requires the Conditional Compile Argument 'Testt = 1'
' ----------------------------------------------------------------

Private Const SECTION_NAME = "Section-" ' for PrivateProfile services test
Private Const VALUE_NAME = "-Name-"     ' for PrivateProfile services test
Private Const VALUE_STRING = "-Value-"  ' for PrivateProfile services test
    
Private cllTestFiles    As Collection

Private Property Get SectionName(Optional ByVal l As Long)
    SectionName = SECTION_NAME & Format(l, "00")
End Property

Private Property Let Status(ByVal s As String)
    If s <> vbNullString Then
        Application.StatusBar = "Regression test " & ThisWorkbook.name & " module 'mFile': " & s
    Else
        Application.StatusBar = vbNullString
    End If
End Property

Private Property Get ValueName(Optional ByVal lS As Long, Optional ByVal lV As Long)
    ValueName = SECTION_NAME & Format(lS, "00") & VALUE_NAME & Format(lV, "00")
End Property

Private Function AppErr(ByVal app_err_no As Long) As Long
' ------------------------------------------------------------------------------
' Ensures that a programmed (i.e. an application) error numbers never conflicts
' with the number of a VB runtime error. Thr function returns a given positive
' number (app_err_no) with the vbObjectError added - which turns it into a
' negative value. When the provided number is negative it returns the original
' positive "application" error number e.g. for being used with an error message.
' ------------------------------------------------------------------------------
    If app_err_no >= 0 Then AppErr = app_err_no + vbObjectError Else AppErr = Abs(app_err_no - vbObjectError)
End Function

Private Property Get ValueString(Optional ByVal lS As Long, Optional ByVal lV As Long)
    ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Property

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = ThisWorkbook.name & ": mTest." & sProc
End Function

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

    mErH.BoTP ErrSrc(PROC), AppErr(1) ' For the very last test on an error condition
    mTest.Regression_Other
    mTest.Regression_PrivateProfile
    
xt: TestFilesRemove
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Regression_Other()
' ---------------------------------------------------------
' All results are asserted and there is no intervention
' required for the whole test. When an assertion fails the
' test procedure will stop and indicates the problem with
' the called procedure.
' An execution trace is displayed for each test procedure.
' ---------------------------------------------------------
    Const PROC = "Regression_Other"

    On Error GoTo eh
    Dim sTestStatus As String
    
    sTestStatus = "mFile Regression-Other: "

    mErH.BoTP ErrSrc(PROC), AppErr(1) ' For the very last test on an error condition
    mTest.Test_00_Temp
    mTest.Test_01_FileExists_Not
    mTest.Test_02_FileExists_ByObject
    mTest.Test_03_FileExists_ByFullName
    mTest.Test_04_FileExists_ByFullName_WildCard_ExactlyOne
    mTest.Test_05_FileExists_ByFullName_WildCard_MoreThanOne
    mTest.Test_06_FileExists_WildCard_MoreThanOne_InSubFolder
    mTest.Test_07_SelectFile
    mTest.Test_08_Txt_Let_Get
    mTest.Test_09_File_Differs
    mTest.Test_10_Arry_Get_Let
    mTest.Test_11_Search
    mTest.Test_99_FileExists_NoFileObject_NoString ' Error AppErr(1) !
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Regression_PrivateProfile()
' ---------------------------------------------------------
' All results are asserted and there is no intervention
' required for the whole test. When an assertion fails the
' test procedure will stop and indicates the problem with
' the called procedure.
' An execution trace is displayed for each test procedure.
' ---------------------------------------------------------
    Const PROC = "Regression_PrivateProfile"

    On Error GoTo eh
    Dim sTestStatus As String
    
    sTestStatus = "mFile Regression_PrivateProfile: "

    mErH.BoTP ErrSrc(PROC), AppErr(1) ' For the very last test on an error condition
    mTest.Test_52_File_Value
    mTest.Test_53_File_Values
    mTest.Test_54_File_ValueNames
    mTest.Test_55_File_SectionNames
    mTest.Test_56_File_PrivateProperty_Exists
    mTest.Test_60_File_SectionsCopy
    
xt: TestFilesRemove
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Sub TestFilesRemove()

    Dim fso As New FileSystemObject
    Dim sFile   As String
    Dim v   As Variant
    
    If cllTestFiles Is Nothing Then Exit Sub
    For Each v In cllTestFiles
        sFile = v
        If fso.FileExists(sFile) Then fso.DeleteFile sFile
    Next v
    
    Set cllTestFiles = Nothing
    Set fso = Nothing
    
End Sub

Private Function TestFileTemp() As String
    Dim sFile   As String
    
    sFile = mFile.Temp(tmp_extension:=".dat")
    TestFileTemp = sFile
    
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFile

End Function

Private Function TestFileWithSections( _
               Optional ByVal ts_section_name As String = "Section-", _
               Optional ByVal ts_value_name As String = "-Name-", _
               Optional ByVal ts_value As String = "-Value-", _
               Optional ByVal ts_sections As Long = 3, _
               Optional ByVal ts_values As Long = 3) As String
' ---------------------------------------------------------------------
' Returns the name of a temporary test file with (ts_sections) sections
' with (ts_values) values each.
' ---------------------------------------------------------------------
    Dim i       As Long
    Dim j       As Long
    Dim sFile   As String
    
    sFile = mFile.Temp(tmp_extension:=".dat")
    For i = ts_sections To 1 Step -1
        For j = ts_values To 1 Step -1
            mFile.Value(pp_file:=sFile _
                      , pp_section:=SectionName(i) _
                      , pp_value_name:=ValueName(i, j) _
                       ) = ValueString(i, j)
        Next j
    Next i
    TestFileWithSections = sFile
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFile
    
End Function

Public Sub Test_00_Temp()
    Const PROC = "Test_00_Temp"

    Dim sTemp As String
    
    mErH.BoP ErrSrc(PROC)
    sTemp = mFile.Temp(tmp_path:=ThisWorkbook.Path)
    sTemp = mFile.Temp()
    mErH.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_01_FileExists_Not()
    Const PROC = "Test_01_FileExists_Not"
    Const NOT_EXIST = "Test.txt"
    
    On Error GoTo eh

    mErH.BoP ErrSrc(PROC), "fe_file:=", NOT_EXIST
    Status = ErrSrc(PROC)
    Debug.Assert mFile.Exists(fe_file:="Test.txt") = False
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    
    mErH.BoP ErrSrc(PROC), "fe_file:=", wb.FullName
    Debug.Assert mFile.Exists(fe_file:=fso) = True
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    mErH.BoP ErrSrc(PROC), "fe_file:=", sWldCrd
    Debug.Assert mFile.Exists(fe_file:=sWldCrd, fe_cll:=cll) = True
    Debug.Assert cll.Count = 1
    Debug.Assert cll.Item(1).Path = fsoFile.Path
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    mErH.BoP ErrSrc(PROC), "fe_file:=", sWldCrd
    Debug.Assert mFile.Exists(fe_file:=sWldCrd, fe_cll:=cllFiles) = True
    Debug.Assert cllFiles.Count = 2
    Debug.Assert cllFiles.Item(1).name = "fMsg.frm"
    Debug.Assert cllFiles.Item(2).name = "fMsg.frx"
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    mErH.BoP ErrSrc(PROC), "fe_file:=", sWldCrd
    Debug.Assert mFile.Exists( _
                              fe_file:=sWldCrd, _
                              fe_cll:=cllFiles _
                             ) = True
    Debug.Assert cllFiles.Count >= 2
    Debug.Assert cllFiles.Item(1).name = "fMsg.frm"
    Debug.Assert cllFiles.Item(2).name = "fMsg.frx"
            
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_08_Txt_Let_Get()
    Const PROC = "Test_08_Txt_Let_Get"
    
    On Error GoTo eh
    Dim sFl     As String
    Dim sTest   As String
    Dim sResult As String
    Dim a()     As String
    Dim sSplit  As String
    Dim fso     As New FileSystemObject
    Dim oFl     As File
    
    Status = ErrSrc(PROC)

    '~~ Test 1: Write one recod
    sFl = mFile.Temp()
    sTest = "My string"
    mFile.Txt(ft_file:=sFl _
            , ft_append:=False _
             ) = sTest
    sResult = mFile.Txt(ft_file:=sFl, ft_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = sTest
    fso.DeleteFile sFl
    
    '~~ Test 2: Empty file
    sFl = mFile.Temp()
    sTest = vbNullString
    mFile.Txt(ft_file:=sFl, ft_append:=False) = sTest
    sResult = mFile.Txt(ft_file:=sFl, ft_split:=sSplit)
    Debug.Assert sResult = vbNullString
    fso.DeleteFile sFl

    '~~ Test 3: Append
    sFl = mFile.Temp()
    mFile.Txt(ft_file:=sFl, ft_append:=False) = "AAA" & vbCrLf & "BBB"
    mFile.Txt(ft_file:=sFl, ft_append:=True) = "CCC"
    sResult = mFile.Txt(ft_file:=sFl, ft_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = "AAA"
    Debug.Assert Split(sResult, sSplit)(1) = "BBB"
    Debug.Assert Split(sResult, sSplit)(2) = "CCC"
    fso.DeleteFile sFl

    '~~ Test 4: Write with append and read with file as object
    sFl = mFile.Temp()
    fso.CreateTextFile Filename:=sFl
    Set oFl = fso.GetFile(sFl)
    sFl = oFl.Path
    mFile.Txt(ft_file:=oFl, ft_append:=False) = "AAA" & vbCrLf & "BBB"
    mFile.Txt(ft_file:=oFl, ft_append:=True) = "CCC"
    sResult = mFile.Txt(ft_file:=oFl, ft_split:=sSplit)
    Debug.Assert Split(sResult, sSplit)(0) = "AAA"
    Debug.Assert Split(sResult, sSplit)(1) = "BBB"
    Debug.Assert Split(sResult, sSplit)(2) = "CCC"
    fso.DeleteFile sFl

xt: Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    mFile.Txt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFile.Txt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)

    ' Test 1: Differs.Count = 0
    mErH.BoP ErrSrc(PROC)
    Set dctDiff = mFile.Differs(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 0

    ' Test 2: Differs.Count = 1
    mFile.Txt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFile.Txt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFile.Differs(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    
    ' Test 3: Differs.Count = 1
    mFile.Txt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C" & vbCrLf & "D"
    mFile.Txt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFile.Differs(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1
    
    ' Test 4: Differs.Count = 1
    mFile.Txt(ft_file:=sF1, ft_append:=False) = "A" & vbCrLf & "B" & vbCrLf & "C"
    mFile.Txt(ft_file:=sF2, ft_append:=False) = "A" & vbCrLf & "X" & vbCrLf & "C"
    Set f1 = fso.GetFile(sF1)
    Set f2 = fso.GetFile(sF2)
    
    Set dctDiff = mFile.Differs(fd_file1:=f1 _
                              , fd_file2:=f2 _
                              , fd_ignore_empty_records:=True _
                              , fd_stop_after:=2 _
                               )
    Debug.Assert dctDiff.Count = 1

xt: mErH.EoP ErrSrc(PROC)
    If fso.FileExists(sF1) Then fso.DeleteFile (sF1)
    If fso.FileExists(sF2) Then fso.DeleteFile (sF2)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
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
    mErH.BoP ErrSrc(PROC), "fd_file1 = ", f1.name, "fd_file2 = ", f2.name
    Set dctDiff = mFile.Differs(fd_file1:=f1, fd_file2:=f2, fd_ignore_empty_records:=True)
    Debug.Assert dctDiff.Count = 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_10_Arry_Get_Let()
    Const PROC = "Test_10_Arry_Get_Let"
    
    On Error GoTo eh
    Dim sFile1      As String
    Dim sFile2      As String
    Dim lInclEmpty  As Long
    Dim lEmpty1     As Long
    Dim lExclEmpty  As Long
    Dim lEmpty2     As Long
    Dim fso         As New FileSystemObject
    Dim a           As Variant
    Dim v           As Variant
    
    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    sFile1 = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas"
    sFile2 = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas"
    
    sFile1 = mFile.Temp()
    sFile2 = mFile.Temp()
    
    '~~ Write to lines to sFile1
    mFile.Txt(sFile1) = "xxx" & vbCrLf & "" & "yyy"
    
    '~~ Get the two lines as Array
    a = mFile.arry(fa_file:=sFile1 _
                 , fa_split:=vbCrLf _
                  )
    Debug.Assert a(LBound(a)) = "xxx"
    Debug.Assert a(UBound(a)) = "yyy"

    '~~ Write array to file-2
    mFile.arry(fa_file:=sFile2 _
             , fa_split:=vbCrLf _
              ) = a
    Debug.Assert mFile.Differs(fso.GetFile(sFile1), fso.GetFile(sFile2)).Count = 0

    '~~ Count empty records when array contains all text lines
    a = mFile.arry(fa_file:=sFile1, fa_excl_empty_lines:=False)
    lInclEmpty = UBound(a) + 1
    lEmpty1 = 0
    For Each v In a
        If VBA.Trim$(v) = vbNullString Then lEmpty1 = lEmpty1 + 1
        If VBA.Len(Trim$(v)) = 0 Then lEmpty2 = lEmpty2 + 1
    Next v
    
    '~~ Count empty records
    a = mFile.arry(fa_file:=sFile1, fa_excl_empty_lines:=True)
    lExclEmpty = UBound(a) + 1
    Debug.Assert lExclEmpty = lInclEmpty - lEmpty1
    
xt: With fso
        .DeleteFile sFile1
        If .FileExists(sFile2) Then .DeleteFile sFile2
    End With
    Set fso = Nothing
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_11_Search()
    Const PROC = "Test_11_Search"
    
    On Error GoTo eh
    Dim cll As Collection
    Dim v   As Variant
    
    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    
    '~~ Test 1: Including subfolders, several files found
    Set cll = mFile.Search(fs_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                         , fs_mask:="*CompMan*.frx" _
                         , fs_stop_after:=5 _
                          )
    Debug.Assert cll.Count > 0

    '~~ Test 2: Not including subfolders, no files found
    Set cll = mFile.Search(fs_root:="e:\Ablage\Excel VBA\DevAndTest\Common" _
                         , fs_mask:="*CompMan*.frx" _
                         , fs_stop_after:=5 _
                         , fs_in_subfolders:=False _
                          )
    Debug.Assert cll.Count = 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_52_File_Value()
' ------------------------------------------------
' This test relies on the Value (Let) service.
' ------------------------------------------------
    Const PROC = "Test_52_File_Value"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sFile       As String
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim cyResult    As Currency
    
    '~~ Test preparation
    sFile = TestFileTemp
        
    mErH.BoP ErrSrc(PROC)
    
    '~~ Test 1: Read non-existing value from a non-existing file
    Debug.Assert mFile.Value(pp_file:=sFile _
                           , pp_section:="Any" _
                           , pp_value_name:="Any" _
                            ) = vbNullString
    
    '~~ Test 2: Write values
    mFile.Value(pp_file:=sFile, pp_section:=SectionName(1), pp_value_name:=ValueName(1, 1)) = ValueString(1, 1)
    mFile.Value(pp_file:=sFile, pp_section:=SectionName(1), pp_value_name:=ValueName(1, 2)) = ValueString(1, 2)
    mFile.Value(pp_file:=sFile, pp_section:=SectionName(2), pp_value_name:=ValueName(2, 1)) = ValueString(2, 1)
    mFile.Value(pp_file:=sFile, pp_section:=SectionName(2), pp_value_name:=ValueName(2, 2)) = cyValue
    
    '~~ Test 2: Assert written values
    Debug.Assert mFile.Value(pp_file:=sFile, pp_section:=SectionName(1), pp_value_name:=ValueName(1, 1)) = ValueString(1, 1)
    Debug.Assert mFile.Value(pp_file:=sFile, pp_section:=SectionName(1), pp_value_name:=ValueName(1, 2)) = ValueString(1, 2)
    Debug.Assert mFile.Value(pp_file:=sFile, pp_section:=SectionName(2), pp_value_name:=ValueName(2, 1)) = ValueString(2, 1)
    cyResult = mFile.Value(pp_file:=sFile, pp_section:=SectionName(2), pp_value_name:=ValueName(2, 2))
    Debug.Assert cyResult = cyValue
    Debug.Assert VarType(cyResult) = vbCurrency
    
xt: mErH.EoP ErrSrc(PROC)
    TestFilesRemove
    Set fso = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_53_File_Values()
    Const PROC = "Test_53_File_Values"
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim v           As Variant
    Dim sFile       As String
    Dim fso         As New FileSystemObject
    Dim sSection    As String
    
    sFile = TestFileWithSections( _
                               , ts_sections:=3 _
                               , ts_values:=3 _
                        )
    
    mErH.BoP ErrSrc(PROC)

    '~~ Test 1: All values of one section
    Set dct = mFile.Values(pp_file:=sFile _
                         , pp_sections:=SectionName(1) _
                          )
    '~~ Test 1: Assert the content of the result Dictionary
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Keys()(0) = ValueString(1, 1)
    Debug.Assert dct.Keys()(1) = ValueString(1, 2)
    Debug.Assert dct.Keys()(2) = ValueString(1, 3)
    '~~ Attention! There may be more value names for a value when the same value appears in several sections under a different value name!
    '~~            Thus, the names are returned as Collection. When the values are from one specific section there will be only one name
    '~~            in the item=collection
    Debug.Assert dct.Items()(0)(1) = ValueName(1, 1)
    Debug.Assert dct.Items()(1)(1) = ValueName(1, 2)
    Debug.Assert dct.Items()(2)(1) = ValueName(1, 3)
    
    '~~ Test 2: All values of all section
    Set dct = mFile.Values(sFile) ' all sections is the default when no name is provided via the pp_sections argument
    Debug.Assert dct.Count = 9

xt: mErH.EoP ErrSrc(PROC)
    TestFilesRemove
    Set dct = Nothing
    Set fso = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_54_File_ValueNames()
    Const PROC = "Test_54_File_ValueNames"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim sFile   As String
    Dim dct     As Dictionary
    Dim fso     As New FileSystemObject
    
    sFile = TestFileWithSections( _
                                 ts_sections:=3 _
                               , ts_values:=3 _
                                )
    
    mErH.BoP ErrSrc(PROC)
    Set dct = mFile.ValueNames(sFile)
    Debug.Assert dct.Count = 9
    
xt: mErH.EoP ErrSrc(PROC)
    TestFilesRemove
    Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_55_File_SectionNames()
    Const PROC = "Test_55_File_SectionNames"
    
    On Error GoTo eh
    Dim v       As Variant
    Dim sFile   As String
    Dim fso     As New FileSystemObject
    Dim dct     As Dictionary
    
    sFile = TestFileWithSections( _
                                 ts_sections:=3 _
                               , ts_values:=3 _
                                )
    
    mErH.BoP ErrSrc(PROC)
    Set dct = mFile.SectionNames(sFile)
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Items()(0) = SectionName(1)
    Debug.Assert dct.Items()(1) = SectionName(2)
    Debug.Assert dct.Items()(2) = SectionName(3)

xt: mErH.EoP ErrSrc(PROC)
    TestFilesRemove
    Set fso = Nothing
    Set dct = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_56_File_PrivateProperty_Exists()
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "Test_56_File_PrivateProperty_Exists"

    On Error GoTo eh
    Dim sFile   As String
    
    '~~ Test preparation
    sFile = TestFileWithSections( _
                               , ts_sections:=10 _
                               , ts_values:=3 _
                                )
    mErH.BoP ErrSrc(PROC)
    '~~ Section by name in file
    Debug.Assert mFile.SectionExists(pp_file:=sFile _
                                  , pp_section:=SectionName(9) _
                                    )
    '~~ Value name in any section
    Debug.Assert mFile.ValueNameExists(pp_file:=sFile _
                                  , pp_valuename:=ValueName(9, 3) _
                                    )
    '~~ Value name in a named section
    Debug.Assert mFile.ValueNameExists(pp_file:=sFile _
                                     , pp_sections:=SectionName(7) _
                                     , pp_valuename:=ValueName(7, 3) _
                                      )
    '~~ Value name not in named section
    Debug.Assert Not mFile.ValueNameExists(pp_file:=sFile _
                                         , pp_sections:=SectionName(7) _
                                         , pp_valuename:=ValueName(6, 3) _
                                          )
    
    '~~ Value in any section
    Debug.Assert mFile.ValueExists(pp_file:=sFile _
                                 , pp_value:=ValueString(8, 3) _
                                  )
    
    '~~ Value in named section
    Debug.Assert mFile.ValueExists(pp_file:=sFile _
                                 , pp_sections:=SectionName(7) _
                                 , pp_value:=ValueString(7, 3) _
                                  )
    '~~ Value not in named section
    Debug.Assert Not mFile.ValueExists(pp_file:=sFile _
                                     , pp_sections:=SectionName(7) _
                                     , pp_value:=ValueString(6, 3) _
                                      )
xt: mErH.EoP ErrSrc(PROC)
    TestFilesRemove
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_60_File_SectionsCopy()
' ------------------------------------------------
' This test relies on successfully tests:
' - Test_55_File_SectionNames (mFile.SectionNames)
' Iplicitely tested are:
' - mFile.Sections Get and Let
' ------------------------------------------------
    Const PROC = "Test_60_File_SectionsCopy"
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim sFileGet        As String
    Dim sFileLet        As String
    Dim i               As Long
    Dim j               As Long
    Dim arSections()    As Variant
    Dim sSectionName    As String
    Dim dct             As Dictionary
    
    '~~ Test preparation
    sFileGet = TestFileWithSections( _
                                    ts_sections:=20 _
                                  , ts_values:=10 _
                                    )
    mErH.BoP ErrSrc(PROC)
    
    '~~ Test 1: Copy the first section only
    sFileLet = mFile.Temp(tmp_extension:=".dat")
    sSectionName = mFile.SectionNames(sFileGet).Items()(0)
    mFile.SectionsCopy pp_source:=sFileGet _
                     , pp_target:=sFileLet _
                     , pp_sections:=sSectionName
    
    '~~ Test 1: Assert result
    Set dct = mFile.SectionNames(sFileLet)
    Debug.Assert dct.Count = 1
    Debug.Assert dct.Keys()(0) = SectionName(1)
    fso.DeleteFile sFileLet
    
    
    '~~ Test 2: Copy all sections
    sFileLet = mFile.Temp(tmp_extension:=".dat")
    mFile.SectionsCopy pp_source:=sFileGet _
                     , pp_target:=sFileLet
    
    '~~ Test 2: Assert result
    Set dct = mFile.SectionNames(sFileLet)
    Debug.Assert dct.Count = 20
    Debug.Assert dct.Keys()(0) = SectionName(1)
    Debug.Assert dct.Keys()(1) = SectionName(2)
    Debug.Assert dct.Keys()(2) = SectionName(3)
    fso.DeleteFile sFileLet
       
    '~~ Test 3: Order sections in ascending sequence
    Debug.Assert mFile.arry(sFileGet)(0) = "[" & SectionName(20) & "]"
    
    mFile.SectionsCopy pp_source:=sFileGet _
                     , pp_target:=sFileGet _
                     , pp_replace:=True ' essential to get them re-ordered in ascending sequence
    
    '~~ Test 3: Assert result
    Debug.Assert mFile.arry(sFileGet)(0) = "[" & SectionName(1) & "]"
            
xt: mErH.EoP ErrSrc(PROC)
    TestFilesRemove
    Set fso = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Public Sub Test_99_FileExists_NoFileObject_NoString()
    Const PROC = "Test_99_FileExists_NoFileObject_NoString"
    
    On Error GoTo eh

    Status = ErrSrc(PROC)
    mErH.BoP ErrSrc(PROC)
    mFile.Exists ThisWorkbook
    Debug.Assert AppErr(1)
        
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbYes: Stop: Resume
        Case vbNo:  Stop: Resume Next
        Case Else:  GoTo xt
    End Select
End Sub

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' This is a kind of universal error message which includes a debugging option.
' It may be copied into any module as a Private Function. The function works
' "standalone" as well with (i.e. uses) the Common VBA Message Component
' (fMsg,mMsg) and with the Common Error Handling Component (ErH) installed.
' Either will be used with the Conditional Compile Argument 'CommMsgComp = 1'
' and/or 'CommErHComp = 1' which provides a better designed error message.
'
' Usage: When this procedure is copied as a Private Function into any desired
'        module an error handling which consideres the possible Conditional
'        Compile Argument 'Debugging = 1' will look as follows
'
'            Const PROC = "procedure-name"
'            On Error Goto eh
'        ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC)
'               Case vbYes: Stop: Resume
'               Case vbNo:  Resume Next
'               Case Else:  Goto xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Used:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
' ------------------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    '~~ Determine the type of error
    Select Case err_no
        Case Is < 0
            ErrNo = AppErr(err_no)
            ErrType = "Application Error "
        Case Else
            ErrNo = err_no
            If (InStr(1, err_dscrptn, "DAO") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC Teradata Driver") <> 0 _
            Or InStr(1, err_dscrptn, "ODBC") <> 0 _
            Or InStr(1, err_dscrptn, "Oracle") <> 0) _
            Then ErrType = "Database Error " _
            Else ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              err_dscrptn & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    
#If Debugging Then
    ErrBttns = vbYesNoCancel
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume error line" & vbLf & _
              "No     = Resume Next (skip error line)" & vbLf & _
              "Cancel = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
#If CommErHComp Then
    '~~ When the Common VBA Error Handling Component (ErH) is installed/used by in the VB-Project
    ErrMsg = mErH.ErrMsg(err_source:=err_source, err_number:=err_no, err_dscrptn:=err_dscrptn, err_line:=err_line)
    '~~ Translate back the elaborated reply buttons mErrH.ErrMsg displays and returns to the simple yes/No/Cancel
    '~~ replies with the VBA MsgBox.
    Select Case ErrMsg
        Case mErH.DebugOptResumeErrorLine:  ErrMsg = vbYes
        Case mErH.DebugOptResumeNext:       ErrMsg = vbNo
        Case Else:                          ErrMsg = vbCancel
    End Select
#Else
    '~~ When the Common VBA Error Handling Component (ErH) is not used/installed there might still be the
    '~~ Common VBA Message Component (Msg) be installed/used
#If CommMsgComp Then
    ErrMsg = mMsg.ErrMsg(err_source:=err_source)
#Else
    '~~ None of the Common Components is installed/used
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
#End If
#End If
End Function


