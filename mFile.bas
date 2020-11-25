Attribute VB_Name = "mFile"
Option Explicit
Option Compare Text
Option Private Module
' --------------------------------------------------------------
' Standard  Module mFile
'           Common methods and functions regarding file objects.
'
' Methods:  Exists      Returns TRUE when the file exists
'           Differ      Returns TRUE when two files have a
'                       different content
'           Delete      Deletes a file
'           Extension   Returns the extension of a file's name
'           GetFile     Returns a file object for a given name
'           ToArray     Returns a file's content in an array
'
' Uses:     No other components (mTrc is for module mTest only).
'
' Requires: Reference to "Microsoft Scripting Runtine"
'
' W. Rauschenberger, Berlin Nov 2020
' -----------------------------------------------------------------------------------
Private Const CONCAT = "||"

Public Function Exists(ByVal exst_file As Variant, _
              Optional ByRef exst_fso As File = Nothing, _
              Optional ByRef exst_cll As Collection = Nothing) As Boolean
' ------------------------------------------------------------------
' Returns TRUE when the file (exst_file) - which may be a file object
' or a file's full name - exists and furthermore:
' - when the file's full name ends with a wildcard * all
'   subfolders are scanned and any file which meets the criteria
'   is returned as File object in a collection (exst_cll),
' - when the files's full name does not end with a wildcard * the
'   existing file is returned as a File object (exst_fso).
' ----------------------------------------------------------------
    Const PROC  As String = "Exists"    ' This procedure's name for the error handling and execution tracking
    
    On Error GoTo eh
    Dim sTest   As String
    Dim sFile   As String
    Dim fldr    As Folder
    Dim sfldr   As Folder   ' Sub-Folder
    Dim fl      As File
    Dim sPath   As String
    Dim queue   As Collection

    Exists = False
    Set exst_cll = New Collection

    If TypeName(exst_file) <> "File" And TypeName(exst_file) <> "String" _
    Then err.Raise AppErr(1), ErrSrc(PROC), "The File (parameter exst_file) for the File's existence check is neither a full path/file name nor a file object!"
    If Not TypeName(exst_fso) = "Nothing" And Not TypeName(exst_fso) = "File" _
    Then err.Raise AppErr(2), ErrSrc(PROC), "The provided return parameter (exst_fso) is not a File type!"
    If Not TypeName(exst_cll) = "Nothing" And Not TypeName(exst_cll) = "Collection" _
    Then err.Raise AppErr(3), ErrSrc(PROC), "The provided return parameter (exst_cll) is not a Collection type!"

    If TypeOf exst_file Is File Then
        With New FileSystemObject
            On Error Resume Next
            sTest = exst_file.Name
            Exists = err.Number = 0
            If Exists Then
                '~~ Return the existing file as File object
                Set exst_fso = .GetFile(exst_file.Path)
                GoTo xt
            End If
        End With
    ElseIf VarType(exst_file) = vbString Then
        With New FileSystemObject
            sFile = Split(exst_file, "\")(UBound(Split(exst_file, "\")))
            If Not Right(sFile, 1) = "*" Then
                Exists = .FileExists(exst_file)
                If Exists Then
                    '~~ Return the existing file as File object
                    Set exst_fso = .GetFile(exst_file)
                    GoTo xt
                End If
            Else
                sPath = Replace(exst_file, "\" & sFile, vbNullString)
                sFile = Replace(sFile, "*", vbNullString)
                '~~ Wildcard file existence check is due
                Set fldr = .GetFolder(sPath)
                Set queue = New Collection
                queue.Add .GetFolder(sPath)

                Do While queue.Count > 0
                    Set fldr = queue(queue.Count)
                    queue.Remove queue.Count ' dequeue the processed subfolder
                    For Each sfldr In fldr.SubFolders
                        queue.Add sfldr ' enqueue (collect) all subfolders
                    Next sfldr
                    For Each fl In fldr.Files
                        If InStr(fl.Name, sFile) <> 0 And left(fl.Name, 1) <> "~" Then
                            '~~ Return the existing file which meets the search criteria
                            '~~ as File object in a collection
                            exst_cll.Add fl
                         End If
                    Next fl
                Loop
                If exst_cll.Count > 0 Then Exists = True
            End If
        End With
    End If

xt: Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function GetFile(ByVal gf_path As String) As File
    With New FileSystemObject
        Set GetFile = .GetFile(gf_path)
    End With
End Function

Public Function ToArray(ByVal ta_file As Variant) As String()
' ---------------------------------------------------------
' Returns the content of the file (vFile) - which may be
' provided as file object or full file name - as array
' by considering any kind of line break characters.
' ---------------------------------------------------------
    Const PROC  As String = "ToArray"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim a       As Variant
    Dim sSplit  As String
    Dim fso     As File
    Dim sFile   As String

    BoP ErrSrc(PROC)
    
    If Not Exists(ta_file, fso) _
    Then err.Raise AppErr(1), ErrSrc(PROC), "The file object (vFile) does not exist!"
    
    '~~ Unload file into a test stream
    With New FileSystemObject
        Set ts = .OpenTextFile(fso.Path, 1)
        With ts
            On Error Resume Next ' may be empty
            sFile = .ReadAll
            .Close
        End With
    End With
    
    If sFile = vbNullString Then GoTo xt
    
    '~~ Get the kind of line break used
    If InStr(sFile, vbCr) <> 0 Then sSplit = vbCr
    If InStr(sFile, vbLf) <> 0 Then sSplit = sSplit & vbLf
    
    '~~ Test stream to array
    a = Split(sFile, sSplit)
    
    '~~ Remove any leading or trailing empty items
    mBasic.ArrayTrimm a
    ToArray = a
    
xt: EoP ErrSrc(PROC)
    Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function ToDict(ByVal td_file As Variant) As Dictionary
' ----------------------------------------------------------
' Returns the content of the file (td_file) - which may be
' provided as file object or full file name - as Dictionary
' by considering any kind of line break characters.
' ---------------------------------------------------------
    Const PROC  As String = "ToDict"
    
    On Error GoTo eh
    Dim ts      As TextStream
    Dim a       As Variant
    Dim dct     As New Dictionary
    Dim sSplit  As String
    Dim fso     As File
    Dim sFile   As String
    Dim i       As Long
    
    If Not Exists(td_file, fso) _
    Then err.Raise AppErr(1), ErrSrc(PROC), "The file object (td_file) does not exist!"
    
    '~~ Unload file into a test stream
    With New FileSystemObject
        Set ts = .OpenTextFile(fso.Path, 1)
        With ts
            On Error Resume Next ' may be empty
            sFile = .ReadAll
            .Close
        End With
    End With
    
    If sFile = vbNullString Then GoTo xt
    
    '~~ Get the kind of line break used
    If InStr(sFile, vbCr) <> 0 Then sSplit = vbCr
    If InStr(sFile, vbLf) <> 0 Then sSplit = sSplit & vbLf
    
    '~~ Test stream to array
    a = Split(sFile, sSplit)
    
    '~~ Remove any leading or trailing empty items
    mBasic.ArrayTrimm a
    
    For i = LBound(a) To UBound(a)
        dct.Add i + 1, a(i)
    Next i
        
xt: Set ToDict = dct
    Exit Function
    
eh: ErrMsg ErrSrc(PROC)
End Function

Public Function SelectFile( _
            Optional ByVal sInitPath As String = vbNullString, _
            Optional ByVal sFilters As String = "*.*", _
            Optional ByVal sFilterName As String = "File", _
            Optional flResult As File) As Boolean
' --------------------------------------------------------------
' When a file had been selected TRUE is returned and the
' selected file is returned as File object (flResult).
' --------------------------------------------------------------

    Dim fDialog As FileDialog
    Dim v       As Variant

    Set fDialog = Application.FileDialog(msoFileDialogFilePicker)
    With fDialog
        .AllowMultiSelect = False
        .Title = "Select a(n) " & sFilterName
        .InitialFileName = sInitPath
        .Filters.Clear
        For Each v In Split(sFilters, ",")
            .Filters.Add sFilterName, v
         Next v
         
        If .Show = -1 Then
            '~~ A fie had been selected
           With New FileSystemObject
            Set flResult = .GetFile(fDialog.SelectedItems(1))
            SelectFile = True
           End With
        End If
        '~~ When no file had been selected the flResult will be Nothing
    End With

End Function

Public Function Differ( _
                 ByVal f1 As File, _
                 ByVal f2 As File, _
        Optional ByVal lStopAfter As Long = 1) As Boolean
' -------------------------------------------------------
' Returns TRUE when the content of file (f1) differs from
' the content in file (f2). The comparison stops after
' (lStopAfter) detected differences. The detected
' different lines are optionally returned (vResult).
' -------------------------------------------------------

    Dim a1      As Variant
    Dim a2      As Variant
    Dim vLines  As Variant

    a1 = mFile.ToArray(f1)
    a2 = mFile.ToArray(f2)
    vLines = mBasic.ArrayCompare(a1, a2, lStopAfter)
    If mBasic.ArrayIsAllocated(vLines) Then
        Differ = True
    End If
    
End Function

Public Function AppErr(ByVal err_no As Long) As Long
' -----------------------------------------------------------------
' Used with Err.Raise AppErr(<l>).
' When the error number <l> is > 0 it is considered an "Application
' Error Number and vbObjectErrror is added to it into a negative
' number in order not to confuse with a VB runtime error.
' When the error number <l> is negative it is considered an
' Application Error and vbObjectError is added to convert it back
' into its origin positive number.
' ------------------------------------------------------------------
    If err_no < 0 Then
        AppErr = err_no - vbObjectError
    Else
        AppErr = vbObjectError + err_no
    End If
End Function

Public Sub Delete(ByVal v As Variant)

    Dim fl  As File

    With New FileSystemObject
        If TypeName(v) = "File" Then
            Set fl = v
            .DeleteFile fl.Path
        ElseIf TypeName(v) = "String" Then
            If .FileExists(v) Then
                .DeleteFile v
            End If
        End If
    End With
    
End Sub

Public Function Extension(ByVal vFile As Variant)

    With New FileSystemObject
        If TypeName(vFile) = "File" Then
            Extension = .GetExtensionName(vFile.Path)
        Else
            Extension = .GetExtensionName(vFile)
        End If
    End With

End Function

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
    mTrc.Finish sTitle
    mTrc.Terminate
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
    ErrSrc = ThisWorkbook.Name & ">mFile" & ">" & sProc
End Function
