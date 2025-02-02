Attribute VB_Name = "Main"
Option Explicit

#If VBA7 Then
    ' 64-bit
    Public Declare PtrSafe Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Public Declare PtrSafe Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' for clsRequest.MakeRequest
#Else
    ' 32-bit
    Public Declare Function URLDownloadToFile Lib "urlmon" Alias "URLDownloadToFileA" (ByVal pCaller As Long, ByVal szURL As String, ByVal szFileName As String, ByVal dwReserved As Long, ByVal lpfnCB As Long) As Long
    Public Declare Sub Sleep Lib "kernel32" (ByVal dwMilliseconds As Long) ' for clsRequest.MakeRequest
#End If

Public Const MsbTitle = "VBA DevTools"

Public oApp As App
Public SelectedApp As String
Public SelItmIdx As Long ' 0-based

Sub OnStart()

    With Application
        .ScreenUpdating = False
        .StatusBar = "Working..."
        .EnableEvents = False
'        .Cursor = xlWait
    End With
    
End Sub

Sub OnExit()
    
    With Application
        .ScreenUpdating = True
        .StatusBar = False
        .EnableEvents = True
        .Cursor = xlDefault
    End With
    
    If Err.Number <> 0 Then
        MsgBox "Error: " & Err.Number & vbCr & vbCr _
            & "Description: " & Err.Description, _
            vbCritical, MsbTitle
        Err.Clear
    End If
    
End Sub

Sub EnsureWorkbook()
' Ensure that at least one workbook is open before opening VDT.
' Do not count
' - extensions: .xlam
' - PERSONAL.XLSB
    
    Dim i As Long
    Dim WbCount As Long
    Dim NameExt As Variant
    Dim wb As Workbook
    
    For Each wb In Workbooks
        NameExt = Split(wb.Name, ".", -1, vbTextCompare)
        If UBound(NameExt) = 0 Then
            ' It's a new book without extension
            WbCount = WbCount + 1
        Else
            If NameExt(1) <> "xlam" And wb.Name <> "PERSONAL.XLSB" Then
                WbCount = WbCount + 1
            End If
        End If
    Next wb
    If WbCount = 0 Then
        Workbooks.Add
    End If
    
End Sub

Function IsWbOpen(ByVal WbName As String) As Boolean
' Check if workbook or add-in is open
    
    Dim Wbook As Workbook
    Dim AddInWb As Variant
    
    If InStr(WbName, ".xlam") > 0 Then
        For Each AddInWb In Application.AddIns2
            If AddInWb.Name = WbName Then
                IsWbOpen = True
                Exit For
            End If
        Next AddInWb
    Else
        For Each Wbook In Workbooks
            If Wbook.Name = WbName Then
                IsWbOpen = True
                Exit For
            End If
        Next Wbook
    End If
    
End Function

Sub MkDirTree(ByVal DirTree As String)
' Make direcotry with subfolders if they don't exist.

' Parameters:
' DirTree - full path. NB! Remote path should always begin with two back-slashes: "\\server123\<and so on>"

    Dim IsRemote As Boolean
    Dim i As Long
    Dim CheckPath As String
    Dim Elems As Variant
    Dim Fso As New Scripting.FileSystemObject
    
    If Not Fso.FolderExists(DirTree) Then
        Elems = Split(DirTree, "\")
        IsRemote = IIf(DirTree Like "\\*", True, False)
        CheckPath = IIf(IsRemote, "\\" & Elems(2) & "\", "")
        With Fso
            For i = LBound(Elems) To UBound(Elems)
                ' Element should not be empty
                ' AND
                ' (path is local OR (path is remote and i > 3))
                ' Because Split uses back-slash,
                ' which creates 2 empty items in Elems array.
                If Elems(i) <> "" _
                        And (Not IsRemote Or (IsRemote And i > 2)) Then
                    CheckPath = CheckPath & Elems(i) & "\"
                    If Not .FolderExists(CheckPath) Then
                        .CreateFolder CheckPath
                    End If
                End If
            Next i
        End With
    End If
End Sub

Sub CopyFileFromTo(ByVal FileFrom As String, ByVal DirTo As String, _
        Optional ByVal WithPostfix As Boolean = False)
    
    Dim PostfixedName As String
    Dim FileTo As String
    Dim NmExt As Variant
    Dim Fso As New Scripting.FileSystemObject
    Dim oFile As Scripting.File
    
    If WithPostfix Then
        ' If the file has never been backed up, then move on
        On Error GoTo SkipFileNotFoundErr
    End If
    Set oFile = Fso.GetFile(FileFrom)
    
    ' Create directory tree if not exists.
    Call MkDirTree(DirTo)
    
    If WithPostfix Then
        NmExt = Split(oFile.Name, ".")
        PostfixedName = NmExt(0) _
            & Format(oFile.DateLastModified, "-yyyymmdd-hhmm.") & NmExt(1)
        FileTo = DirTo & "\" & PostfixedName
    Else
        FileTo = DirTo & "\" & oFile.Name
    End If
    
    oFile.Copy Destination:=FileTo, OverWriteFiles:=True
SkipFileNotFoundErr:
    If Err.Number = 53 Then
        MsgBox "This is the first time the file is backed up to BackupDir. " _
            & "Nothing will be added to older versions directory.", _
            vbInformation, MsbTitle
        Err.Clear
    End If
    
End Sub

Sub JoinSelectionToString()
' Join selection's cell values into a string using Sep(arator).
    
    Const Sep As String = ";"
    
    Dim Result As String
    Dim Cell As Range
    
    For Each Cell In Selection
        Result = Result & Cell.Value & Sep
    Next Cell
    Result = Left(Result, Len(Result) - 1)
    Debug.Print Result
    
End Sub

Sub SaveStringToFile(ByVal PrintText As String, Optional ByVal ToOpen As Boolean = False)
' Print string to file.
    
    Dim FNum As Integer
    Dim FPath As String
    
    FPath = Environ("tmp") & "\VBADevTools_OUTPUT.txt"
    FNum = FreeFile
    Open FPath For Output As FNum
    Print #FNum, PrintText
    Close #FNum
    
    If ToOpen Then
        Shell """C:\Program Files\Notepad++\notepad++.exe"" """ & FPath & """", vbNormalFocus
    End If
    
End Sub
