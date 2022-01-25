Attribute VB_Name = "mFileTest"
Option Explicit
' ----------------------------------------------------------------
' Standard Module mFileTest: Test of all services of the module.
'
' ----------------------------------------------------------------
Private Const SECTION_NAME = "Section-" ' for PrivateProfile services test
Private Const VALUE_NAME = "-Name-"     ' for PrivateProfile services test
Private Const VALUE_STRING = "-Value-"  ' for PrivateProfile services test
    
Private cllTestFiles    As Collection

Private Property Get TestProc_SectionName(Optional ByVal l As Long)
    TestProc_SectionName = SECTION_NAME & Format(l, "00")
End Property

Private Property Let Status(ByVal s As String)
    If s <> vbNullString Then
        Application.StatusBar = "Regression test " & ThisWorkbook.Name & " module 'mFile': " & s
    Else
        Application.StatusBar = vbNullString
    End If
End Property

Private Property Get TestProc_ValueName(Optional ByVal lS As Long, Optional ByVal lV As Long)
    TestProc_ValueName = SECTION_NAME & Format(lS, "00") & VALUE_NAME & Format(lV, "00")
End Property

Private Property Get TestProc_ValueString(Optional ByVal lS As Long, Optional ByVal lV As Long)
    TestProc_ValueString = SECTION_NAME & Format(lS, "00") & VALUE_STRING & Format(lV, "00")
End Property

Private Function ErrMsg(ByVal err_source As String, _
               Optional ByVal err_no As Long = 0, _
               Optional ByVal err_dscrptn As String = vbNullString, _
               Optional ByVal err_line As Long = 0) As Variant
' ------------------------------------------------------------------------------
' Universal error message display service including a debugging option active
' when the Conditional Compile Argument 'Debugging = 1' and an optional
' additional "About the error:" section displaying text connected to an error
' message by two vertical bars (||).
'
' A copy of this function is used in each procedure with an error handling
' (On error Goto eh).
'
' The function considers the Common VBA Error Handling Component (ErH) which
' may be installed (Conditional Compile Argument 'ErHComp = 1') and/or the
' Common VBA Message Display Component (mMsg) installed (Conditional Compile
' Argument 'MsgComp = 1'). Only when none of the two is installed the error
' message is displayed by means of the VBA.MsgBox.
'
' Usage: Example with the Conditional Compile Argument 'Debugging = 1'
'
'        Private/Public <procedure-name>
'            Const PROC = "<procedure-name>"
'
'            On Error Goto eh
'            ....
'        xt: Exit Sub/Function/Property
'
'        eh: Select Case ErrMsg(ErrSrc(PROC))
'               Case vbResume:  Stop: Resume
'               Case Else:      GoTo xt
'            End Select
'        End Sub/Function/Property
'
'        The above may appear a lot of code lines but will be a godsend in case
'        of an error!
'
' Uses:  - For programmed application errors (Err.Raise AppErr(n), ....) the
'          function AppErr will be used which turns the positive number into a
'          negative one. The error message will regard a negative error number
'          as an 'Application Error' and will use AppErr to turn it back for
'          the message into its original positive number. Together with the
'          ErrSrc there will be no need to maintain numerous different error
'          numbers for a VB-Project.
'        - The caller provides the source of the error through the module
'          specific function ErrSrc(PROC) which adds the module name to the
'          procedure name.
'
' W. Rauschenberger Berlin, Nov 2021
' ------------------------------------------------------------------------------
#If ErHComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When the Common VBA Error Handling Component (mErH) is installed in the
    '~~ VB-Project (which includes the mMsg component) the mErh.ErrMsg service
    '~~ is preferred since it provides some enhanced features like a path to the
    '~~ error.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mErH.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#ElseIf MsgComp = 1 Then
    '~~ ------------------------------------------------------------------------
    '~~ When only the Common Message Services Component (mMsg) is installed but
    '~~ not the mErH component the mMsg.ErrMsg service is preferred since it
    '~~ provides an enhanced layout and other features.
    '~~ ------------------------------------------------------------------------
    ErrMsg = mMsg.ErrMsg(err_source, err_no, err_dscrptn, err_line)
    GoTo xt
#End If
    '~~ -------------------------------------------------------------------
    '~~ When neither the mMsg nor the mErH component is installed the error
    '~~ message is displayed by means of the VBA.MsgBox
    '~~ -------------------------------------------------------------------
    Dim ErrBttns    As Variant
    Dim ErrAtLine   As String
    Dim ErrDesc     As String
    Dim ErrLine     As Long
    Dim ErrNo       As Long
    Dim ErrSrc      As String
    Dim ErrText     As String
    Dim ErrTitle    As String
    Dim ErrType     As String
    Dim ErrAbout    As String
        
    '~~ Obtain error information from the Err object for any argument not provided
    If err_no = 0 Then err_no = Err.Number
    If err_line = 0 Then ErrLine = Erl
    If err_source = vbNullString Then err_source = Err.Source
    If err_dscrptn = vbNullString Then err_dscrptn = Err.Description
    If err_dscrptn = vbNullString Then err_dscrptn = "--- No error description available ---"
    
    If InStr(err_dscrptn, "||") <> 0 Then
        ErrDesc = Split(err_dscrptn, "||")(0)
        ErrAbout = Split(err_dscrptn, "||")(1)
    Else
        ErrDesc = err_dscrptn
    End If
    
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
            Else: ErrType = "VB Runtime Error "
    End Select
    
    If err_source <> vbNullString Then ErrSrc = " in: """ & err_source & """"   ' assemble ErrSrc from available information"
    If err_line <> 0 Then ErrAtLine = " at line " & err_line                    ' assemble ErrAtLine from available information
    ErrTitle = Replace(ErrType & ErrNo & ErrSrc & ErrAtLine, "  ", " ")         ' assemble ErrTitle from available information
       
    ErrText = "Error: " & vbLf & _
              ErrDesc & vbLf & vbLf & _
              "Source: " & vbLf & _
              err_source & ErrAtLine
    If ErrAbout <> vbNullString _
    Then ErrText = ErrText & vbLf & vbLf & _
                  "About: " & vbLf & _
                  ErrAbout
    
#If Debugging Then
    ErrBttns = vbYesNo
    ErrText = ErrText & vbLf & vbLf & _
              "Debugging:" & vbLf & _
              "Yes    = Resume Error Line" & vbLf & _
              "No     = Terminate"
#Else
    ErrBttns = vbCritical
#End If
    
    ErrMsg = MsgBox(Title:=ErrTitle _
                  , Prompt:=ErrText _
                  , Buttons:=ErrBttns)
xt: Exit Function

End Function

Private Function ErrSrc(ByVal sProc As String) As String
    ErrSrc = "mFileTest." & sProc
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
    mFileTest.Regression_Other
    mFileTest.Regression_PrivateProfile
    
xt: TestProc_RemoveTestFiles
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
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
    mFileTest.Test_01_File_Temp
    mFileTest.Test_02_File_Exists
    mFileTest.Test_07_File_Select
    mFileTest.Test_08_File_Txt_Let_Get
    mFileTest.Test_09_File_Differs
    mFileTest.Test_10_File_Arry_Get_Let
    mFileTest.Test_11_File_Search
    
xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Regression_PrivateProfile()
' ----------------------------------------------------------------------------
' All results are asserted and there is no intervention required for the whole
' test. When an assertion fails the test procedure will stop and indicate the
' problem with the called procedure. An execution trace is displayed for each
' test procedure.
' ----------------------------------------------------------------------------
    Const PROC = "Regression_PrivateProfile"

    On Error GoTo eh
    Dim sTestStatus As String
    
    sTestStatus = "mFile Regression_PrivateProfile: "
    mErH.BoTP ErrSrc(PROC), AppErr(1) ' For the very last test on an error condition
    
    mFileTest.Test_92_PrivateProfile_File_Value
    mFileTest.Test_93_PrivateProfile_File_Values
    mFileTest.Test_94_PrivateProfile_File_ValueNames
    mFileTest.Test_95_PrivateProfile_File_SectionsNames
    mFileTest.Test_96_PrivateProfile_File_Items_Exists
    mFileTest.Test_97_PrivateProfile_File_SectionsCopy
    
xt: TestProc_RemoveTestFiles
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Private Sub TestProc_RemoveTestFiles()

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

Private Function TestProc_TempFile() As String
    Dim sFile   As String
    
    sFile = mFile.Temp(tmp_extension:=".dat")
    TestProc_TempFile = sFile
    
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFile

End Function

Private Function TestProc_PrivateProfile_File(ByVal ts_sections As Long, _
                                              ByVal ts_values As Long) As String
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
            Debug.Print "TestProc_ValueString(i, j) = " & TestProc_ValueString(i, j)
            mFile.Value(pp_file:=sFile _
                      , pp_section:=TestProc_SectionName(i) _
                      , pp_value_name:=TestProc_ValueName(i, j) _
                       ) = TestProc_ValueString(i, j)
        Next j
    Next i
    TestProc_PrivateProfile_File = sFile
    If cllTestFiles Is Nothing Then Set cllTestFiles = New Collection
    cllTestFiles.Add sFile
    
End Function

Public Sub Test_01_File_Temp()
    Const PROC = "Test_01_File_Temp"

    Dim sTemp As String
    
    mErH.BoP ErrSrc(PROC)
    sTemp = mFile.Temp(tmp_path:=ThisWorkbook.Path)
    sTemp = mFile.Temp()
    mErH.EoP ErrSrc(PROC)
    
End Sub

Public Sub Test_02_File_Exists()
    Const PROC = "Test_02_File_Exists"
    
    On Error GoTo eh
    Dim cll     As Collection
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    
    mErH.BoP ErrSrc(PROC)
    
    '~~ Folder exists
    Debug.Assert mFile.Exists(ex_folder:=ThisWorkbook.Path & "x") = False
    Debug.Assert mFile.Exists(ex_folder:=ThisWorkbook.Path) = True
    
    '~~ File exists
    Debug.Assert mFile.Exists(ex_folder:=ThisWorkbook.Path, ex_file:=ThisWorkbook.Name & "x") = False
    Debug.Assert mFile.Exists(ex_folder:=ThisWorkbook.Path, ex_file:=ThisWorkbook.Name) = True

    '~~ Section exists
    sFile = TestProc_PrivateProfile_File(ts_sections:=3, ts_values:=3)
    Debug.Assert mFile.Exists(ex_folder:=fso.GetFile(sFile).ParentFolder _
                            , ex_file:=fso.GetFile(sFile).Name _
                            , ex_section:=TestProc_SectionName(2) _
                             ) = True
    
    '~~ Value-Name exists
    Debug.Assert mFile.Exists(ex_folder:=fso.GetFile(sFile).ParentFolder _
                            , ex_file:=fso.GetFile(sFile).Name _
                            , ex_section:=TestProc_SectionName(2) _
                            , ex_value_name:=TestProc_ValueName(2, 2) _
                             ) = True

    '~~ File by wildcard, in any sub-folder, exactly one
    Debug.Assert mFile.Exists(ex_folder:=ThisWorkbook.Path _
                            , ex_file:=Replace(ThisWorkbook.Name, "xlsb", "xl*") _
                            , ex_result_files:=cll) = True
    Debug.Assert cll.Count = 1
    Debug.Assert cll(1).Path = ThisWorkbook.FullName
            
    '~~ File by wildcard, in any sub-folder, more than one
    Debug.Assert mFile.Exists(ex_folder:=ThisWorkbook.Path, ex_file:="fMsg.fr*", ex_result_files:=cll) = True
    Debug.Assert cll.Count = 2
    Debug.Assert cll(1).Name = "fMsg.frm"
    Debug.Assert cll(2).Name = "fMsg.frx"
                        
xt: TestProc_RemoveTestFiles
    mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_07_File_Select()
    Const PROC = "Test_07_File_Select"
    
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_08_File_Txt_Let_Get()
    Const PROC = "Test_08_File_Txt_Let_Get"
    
    On Error GoTo eh
    Dim sFl     As String
    Dim sTest   As String
    Dim sResult As String
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_09_File_Differs()
    Const PROC = "Test_09_File_Differs"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim f1      As File
    Dim f2      As File
    Dim dctDiff As Dictionary
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

  
Public Sub Test_09_File_Differs_False()
    Const PROC = "Test_09_File_Differs"
    
    On Error GoTo eh
    Dim fso     As New FileSystemObject
    Dim sFile   As String
    Dim f1      As File
    Dim f2      As File
    Dim dctDiff As Dictionary
    
    Status = ErrSrc(PROC)
    ' Prepare
    sFile = "E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas"
    Set f1 = fso.GetFile("E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas")
    Set f2 = fso.GetFile("E:\Ablage\Excel VBA\DevAndTest\Common\File\mFile.bas")
    
    ' Test
    mErH.BoP ErrSrc(PROC), "fd_file1 = ", f1.Name, "fd_file2 = ", f2.Name
    Set dctDiff = mFile.Differs(fd_file1:=f1, fd_file2:=f2, fd_ignore_empty_records:=True)
    Debug.Assert dctDiff.Count = 0

xt: mErH.EoP ErrSrc(PROC)
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

  
Public Sub Test_10_File_Arry_Get_Let()
    Const PROC = "Test_10_File_Arry_Get_Let"
    
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
    a = mFile.Arry(fa_file:=sFile1 _
                 , fa_split:=vbCrLf _
                  )
    Debug.Assert a(LBound(a)) = "xxx"
    Debug.Assert a(UBound(a)) = "yyy"

    '~~ Write array to file-2
    mFile.Arry(fa_file:=sFile2 _
             , fa_split:=vbCrLf _
              ) = a
    Debug.Assert mFile.Differs(fso.GetFile(sFile1), fso.GetFile(sFile2)).Count = 0

    '~~ Count empty records when array contains all text lines
    a = mFile.Arry(fa_file:=sFile1, fa_excl_empty_lines:=False)
    lInclEmpty = UBound(a) + 1
    lEmpty1 = 0
    For Each v In a
        If VBA.Trim$(v) = vbNullString Then lEmpty1 = lEmpty1 + 1
        If VBA.Len(Trim$(v)) = 0 Then lEmpty2 = lEmpty2 + 1
    Next v
    
    '~~ Count empty records
    a = mFile.Arry(fa_file:=sFile1, fa_excl_empty_lines:=True)
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_11_File_Search()
    Const PROC = "Test_11_File_Search"
    
    On Error GoTo eh
    Dim cll As Collection
    
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
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_92_PrivateProfile_File_Value()
' ------------------------------------------------
' This test relies on the Value (Let) service.
' ------------------------------------------------
    Const PROC = "Test_92_PrivateProfile_File_Value"
    
    On Error GoTo eh
    Dim fso         As New FileSystemObject
    Dim sFile       As String
    Dim cyValue     As Currency: cyValue = 12345.6789
    Dim cyResult    As Currency
    
    '~~ Test preparation
    sFile = TestProc_TempFile
        
    mErH.BoP ErrSrc(PROC)
    
    '~~ Test 1: Read non-existing value from a non-existing file
    Debug.Assert mFile.Value(pp_file:=sFile _
                           , pp_section:="Any" _
                           , pp_value_name:="Any" _
                            ) = vbNullString
    
    '~~ Test 2: Write values
    mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 1)) = TestProc_ValueString(1, 1)
    mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 2)) = TestProc_ValueString(1, 2)
    mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 1)) = TestProc_ValueString(2, 1)
    mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 2)) = cyValue
    
    '~~ Test 2: Assert written values
    Debug.Assert mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 1)) = TestProc_ValueString(1, 1)
    Debug.Assert mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(1), pp_value_name:=TestProc_ValueName(1, 2)) = TestProc_ValueString(1, 2)
    Debug.Assert mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 1)) = TestProc_ValueString(2, 1)
    cyResult = mFile.Value(pp_file:=sFile, pp_section:=TestProc_SectionName(2), pp_value_name:=TestProc_ValueName(2, 2))
    Debug.Assert cyResult = cyValue
    Debug.Assert VarType(cyResult) = vbCurrency
    
xt: mErH.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set fso = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_93_PrivateProfile_File_Values()
    Const PROC = "Test_93_PrivateProfile_File_Values"
    On Error GoTo eh
    Dim dct         As Dictionary
    Dim sFile       As String
    Dim fso         As New FileSystemObject
    
    sFile = TestProc_PrivateProfile_File(ts_sections:=3 _
                               , ts_values:=3 _
                                )
    
    mErH.BoP ErrSrc(PROC)

    '~~ Test 1: All values of one section
    Set dct = mFile.Values(pp_file:=sFile _
                         , pp_section:=TestProc_SectionName(1) _
                          )
    '~~ Test 1: Assert the content of the result Dictionary
    Debug.Assert dct.Count = 3
    Debug.Print dct.Keys()(0)
    Debug.Print dct.Keys()(1)
    Debug.Print dct.Keys()(2)
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(1, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(1, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(1, 3)
    '~~ Attention! There may be more value names for a value when the same value appears in several sections under a different value name!
    '~~            Thus, the names are returned as Collection. When the values are from one specific section there will be only one name
    '~~            in the item=collection
    Debug.Assert dct.Items()(0)(1) = TestProc_ValueString(1, 1)
    Debug.Assert dct.Items()(1)(1) = TestProc_ValueString(1, 2)
    Debug.Assert dct.Items()(2)(1) = TestProc_ValueString(1, 3)
    
    '~~ Test 2: No section provided
    Debug.Assert mFile.Values(sFile, vbNullString).Count = 0

    '~~ Test 3: Section does not exist
    Debug.Assert mFile.Values(sFile, "xxxxxxx").Count = 0

xt: mErH.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set dct = Nothing
    Set fso = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_94_PrivateProfile_File_ValueNames()
    Const PROC = "Test_94_PrivateProfile_File_ValueNames"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim dct     As Dictionary
    Dim fso     As New FileSystemObject
    
    sFile = TestProc_PrivateProfile_File(ts_sections:=5, ts_values:=3)
    
    mErH.BoP ErrSrc(PROC)
    Set dct = mFile.ValueNames(sFile)
    Debug.Assert dct.Count = 9
    Debug.Assert dct.Keys()(0) = TestProc_ValueName(1, 1)
    Debug.Assert dct.Keys()(1) = TestProc_ValueName(1, 2)
    Debug.Assert dct.Keys()(2) = TestProc_ValueName(1, 3)
    
    
xt: mErH.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set fso = Nothing
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_95_PrivateProfile_File_SectionsNames()
    Const PROC = "Test_95_PrivateProfile_File_SectionsNames"
    
    On Error GoTo eh
    Dim sFile   As String
    Dim fso     As New FileSystemObject
    Dim dct     As Dictionary
    
    sFile = TestProc_PrivateProfile_File(ts_sections:=3, ts_values:=3)
    
    mErH.BoP ErrSrc(PROC)
    Set dct = mFile.SectionNames(sFile)
    Debug.Assert dct.Count = 3
    Debug.Assert dct.Items()(0) = TestProc_SectionName(1)
    Debug.Assert dct.Items()(1) = TestProc_SectionName(2)
    Debug.Assert dct.Items()(2) = TestProc_SectionName(3)

xt: mErH.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set fso = Nothing
    Set dct = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_96_PrivateProfile_File_Items_Exists()
' ----------------------------------------------------
'
' ----------------------------------------------------
    Const PROC = "Test_96_PrivateProfile_File_Items_Exists"

    On Error GoTo eh
    Dim sFile   As String
    
    '~~ Test preparation
    sFile = TestProc_PrivateProfile_File(ts_sections:=10, ts_values:=3)
    
    mErH.BoP ErrSrc(PROC)
    '~~ Section by name in file
    Debug.Assert mFile.SectionExists(pp_file:=sFile _
                                   , pp_section:=TestProc_SectionName(9) _
                                    )
    '~~ Value name in any section
    Debug.Assert mFile.ValueNameExists(pp_file:=sFile _
                                     , pp_valuename:=TestProc_ValueName(9, 3) _
                                    )
    '~~ Value name in a named section
    Debug.Assert mFile.ValueNameExists(pp_file:=sFile _
                                     , pp_sections:=TestProc_SectionName(7) _
                                     , pp_valuename:=TestProc_ValueName(7, 3) _
                                      )
    '~~ Value name not in named section
    Debug.Assert Not mFile.ValueNameExists(pp_file:=sFile _
                                         , pp_sections:=TestProc_SectionName(7) _
                                         , pp_valuename:=TestProc_ValueName(6, 3) _
                                          )
    
    '~~ Value in any section
    Debug.Assert mFile.ValueExists(pp_file:=sFile _
                                 , pp_value:=TestProc_ValueString(8, 3) _
                                  )
    
    '~~ Value in named section
    Debug.Assert mFile.ValueExists(pp_file:=sFile _
                                 , pp_sections:=TestProc_SectionName(7) _
                                 , pp_value:=TestProc_ValueString(7, 3) _
                                  )
    '~~ Value not in named section
    Debug.Assert Not mFile.ValueExists(pp_file:=sFile _
                                     , pp_sections:=TestProc_SectionName(7) _
                                     , pp_value:=TestProc_ValueString(6, 3) _
                                      )
xt: mErH.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Exit Sub

eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

Public Sub Test_97_PrivateProfile_File_SectionsCopy()
' ------------------------------------------------
' This test relies on successfully tests:
' - Test_95_PrivateProfile_File_SectionsNames (mFile.SectionNames)
' Iplicitely tested are:
' - mFile.Sections Get and Let
' ------------------------------------------------
    Const PROC = "Test_97_PrivateProfile_File_SectionsCopy"
    
    On Error GoTo eh
    Dim fso             As New FileSystemObject
    Dim sFileGet        As String
    Dim sFileLet        As String
    Dim sSectionName    As String
    Dim dct             As Dictionary
    
    '~~ Test preparation
    sFileGet = TestProc_PrivateProfile_File(ts_sections:=20, ts_values:=10)
    
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
    Debug.Assert dct.Keys()(0) = TestProc_SectionName(1)
    fso.DeleteFile sFileLet
    
    
    '~~ Test 2: Copy all sections
    sFileLet = mFile.Temp(tmp_extension:=".dat")
    mFile.SectionsCopy pp_source:=sFileGet _
                     , pp_target:=sFileLet
    
    '~~ Test 2: Assert result
    Set dct = mFile.SectionNames(sFileLet)
    Debug.Assert dct.Count = 20
    Debug.Assert dct.Keys()(0) = TestProc_SectionName(1)
    Debug.Assert dct.Keys()(1) = TestProc_SectionName(2)
    Debug.Assert dct.Keys()(2) = TestProc_SectionName(3)
    fso.DeleteFile sFileLet
       
    '~~ Test 3: Order sections in ascending sequence
    Debug.Assert mFile.Arry(sFileGet)(0) = "[" & TestProc_SectionName(20) & "]"
    
    mFile.SectionsCopy pp_source:=sFileGet _
                     , pp_target:=sFileGet _
                     , pp_replace:=True ' essential to get them re-ordered in ascending sequence
    
    '~~ Test 3: Assert result
    Debug.Assert mFile.Arry(sFileGet)(0) = "[" & TestProc_SectionName(1) & "]"
            
xt: mErH.EoP ErrSrc(PROC)
    TestProc_RemoveTestFiles
    Set fso = Nothing
    Exit Sub
    
eh: Select Case ErrMsg(ErrSrc(PROC))
        Case vbResume:  Stop: Resume
        Case Else:      GoTo xt
    End Select
End Sub

