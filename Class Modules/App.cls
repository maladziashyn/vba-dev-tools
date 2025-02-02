VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "App"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

' Add references:
' Microsoft Visual Basic for Applications Extensibility 5.3. E.g. VBIDE.VBComponents
' Microsoft Scripting Runtime. E.g. Dictionary
' Microsoft VBScript Regular Expressions 5.5. E.g. RegExp

Private Const dirModules As String = "\Modules" ' 1 - vbext_ct_StdModule
Private Const dirClassModules As String = "\Class Modules" ' 2 - vbext_ct_ClassModule
Private Const dirObjects As String = "\Microsoft Excel Objects" ' 100 - vbext_ct_Document: ThisWorkbook, sheets
Private Const dirForms As String = "\Forms" ' 3 - vbext_ct_MSForm
Private Const dirOlderVer As String = "\older-versions"
Private Const dirBuilds As String = "\Builds"

Private Type TApp
    RequiresCodeImport As Boolean
    RequiresEnableEvents As Boolean
    IsAppForBuild As Boolean
    IsAppForDeployTest As Boolean
    IsAppForDeployProd As Boolean

    IsOpen As Boolean

    AppCell As Range ' app's cell in tblApps
    AppFile As String ' e.g.: app.xlam
    AppName As String ' e.g.: app
    AppFPath As String ' e.g.: ...\CodeDir\app.xlam
    AppBackupFullName As String
'    AppRemoteFullName As String
    AppGitRepoURL As String
    AppDeployFName As String ' file name that is sent to server on build
    AppPwd As String
    InjectBuild As String ' stored on wsMain as "sheet index; row; column", e.g. "1;1;1"
    LauncherSub As String ' e.g." StartUpApp

    ' Paths with "\" - backslashes, not for Git commands
    CodeDir As String
    BackupDir As String
    RemoteDir As String
    
    EditXlsm As Boolean

    BuildBackupTestDir As String
    BuildBackupProdDir As String

    ' Path with "/" - forwardslash, for Git commands
    ' when bash does "cd" to this directory
    BashDir As String

    wb As Workbook

    VBComps As VBIDE.VBComponents
    dCurrentVBComps As New Dictionary
    dFilesInCodeDir As New Dictionary
    arrVBCompDirs As Variant
End Type
Private self As TApp

Private Sub Class_Initialize()
    
    Dim arrTmp As Variant
    
' CHECK: App should be on the list
    
    With self
        Set .AppCell = wsMain.ListObjects("tblApps") _
            .ListColumns("App full name [AppCell]") _
            .DataBodyRange.Find(What:=SelectedApp)
        If .AppCell Is Nothing Then
            Err.Raise 800, "App_Initialize", _
                "App is not on the list."
        End If
        
        .RequiresCodeImport = .AppCell.Offset(0, 1)
        .LauncherSub = .AppCell.Offset(0, 2)
        .RequiresEnableEvents = .AppCell.Offset(0, 3)
        .IsAppForBuild = .AppCell.Offset(0, 4)
        .IsAppForDeployTest = .AppCell.Offset(0, 5)
        .IsAppForDeployProd = .AppCell.Offset(0, 6)
        
        .AppDeployFName = .AppCell.Offset(0, 7)
        .AppPwd = .AppCell.Offset(0, 8)
        .InjectBuild = .AppCell.Offset(0, 9)
        
        .AppFile = SelectedApp
        .IsOpen = IsWbOpen(.AppFile)
        
        arrTmp = Split(.AppFile, ".")
        .AppName = arrTmp(0)
        .CodeDir = .AppCell.Offset(0, 10)
        .AppFPath = .CodeDir & "\" & .AppFile
        
        .BackupDir = .AppCell.Offset(0, 11)
        
        Call MkDirTree(.BackupDir)
        .AppBackupFullName = .BackupDir & "\" & .AppFile
        
        .EditXlsm = .AppCell.Offset(0, 12).Value
        
'        .RemoteDir = ServerTest & "path\to\apps\dir" & .AppName ' ?
'        .AppRemoteFullName = .RemoteDir & "\" & .AppFile ' ?
'
'        .AppGitRepoURL = "http://1.2.3.4/Bonobo.Git.Server/" & .AppName & ".git" ' ?
'        .BuildBackupTestDir = wsMain.Range("SharedDir").Value & "\BuildBackups\Test\" & .AppName ' ?
'        .BuildBackupProdDir = wsMain.Range("SharedDir").Value & "\BuildBackups\Prod\" & .AppName ' ?
'
'        Call MkDirTree(.BuildBackupTestDir) ' ?
'        Call MkDirTree(.BuildBackupProdDir) ' ?
        
        ' Git bash needs forward-slashes to "cd"!
        .BashDir = Replace(.CodeDir, "\", "/")
        
        If .IsOpen Then
            Call SetVBComponents
        End If
        
    End With
    
End Sub

Private Sub SetVBComponents()
' Set wb and VB components.
' Create dictionary of existing VB components.
' Necessary for successfully importing the code.
    
    Dim VBComp As VBIDE.VBComponent
    
    Set self.wb = Workbooks(self.AppFile)
    ' VB component dirs, like "...\CodeDir\Modules", etc
    Set self.VBComps = self.wb.VBProject.VBComponents
    self.arrVBCompDirs = Array(dirObjects, dirForms, dirModules, dirClassModules)
    
    ' Make dictionary of existing VB components
    Set self.dCurrentVBComps = New Dictionary
    For Each VBComp In self.VBComps
        self.dCurrentVBComps.Add VBComp.Name, Nothing
    Next VBComp
    
End Sub

Sub AppOpenForEditing(Optional ByVal ForBuild As Boolean = False)
' Open app's workbook taking into account app's attributes:
' if requires code import, if EnableEvents = True, if has launcher sub,
' if has password.
    
    Dim WbOpen As Workbook
    
    If IsWbOpen(self.AppFile) Then
        Err.Raise 850, "OpenForEditing", self.AppFile & " is already open."
    End If
    
    Application.EnableEvents = self.RequiresEnableEvents
    Set WbOpen = Workbooks.Open(self.AppFPath)
    If self.AppPwd <> "" _
            And Application _
                .Workbooks(self.AppFile) _
                .VBProject _
                .Protection = vbext_pp_locked Then
        LockVBProject(WorkbookName:=self.AppFile, password:=self.AppPwd) = False
    End If
    DoEvents
    If self.RequiresCodeImport Then
        Call SetVBComponents
        Call CodeImport(ShowMsbox:=False)
    End If

    ' To run launcher sub, EnableEvents should be set to True
    If Not ForBuild Then
        With Application
            If Not .EnableEvents Then
                .EnableEvents = True
            End If
        End With
        If self.LauncherSub <> "" Then
            Application.Run "'" & self.AppFile & "'!" & self.LauncherSub
        End If
        self.IsOpen = IsWbOpen(self.AppFile)
        RbxUI_VDT.Invalidate
    End If

End Sub

Sub CodeImport(Optional ByVal CheckExisting As Boolean = False, _
        Optional ByVal ShowMsbox As Boolean = True)
' Remove existing modules, classes, forms, and replace with the ones from CodeDir.
' Deletion goes in sub ImportComponent.
' Exception: worksheets and ThisWorkbook,
' only text is updated instead, by deleting existing and pasting new.
' No worksheet or ThisWorkbook is removed.
' If current VBProject has components not present in local Git folder,
' the developer gets the message box and listing of the new components
' in Immediate window.
'
' SPECIAL NOTICE for importing code for "VBA DevTools":
' After import, remove "1" from "App1", etc.
    
    Dim i As Long
    Dim ThisDir As String
    Dim FileNm As String
    Dim VBCdir As Variant
    Dim Fso As New Scripting.FileSystemObject
    Dim oFolder As Scripting.Folder
    Dim oFile As Scripting.File
    Dim oFiles As Scripting.Files
    Dim VBComp As VBIDE.VBComponent
    Dim CurrentVBComps As New Dictionary
    Dim NewVBComps As New Dictionary
    Dim reWsModule As New RegExp
    
    ' This RegEx helps remove first couple lines of code form module file (exported)
    ' that are not visible in VBEditor.
    With reWsModule
        .Pattern = "[\s\w\W\t\r\n)]*Attribute VB_Exposed = (True|False)[\r\n]*"
        .Global = False
        .MultiLine = True
    End With
    
    ' Loop over component directories
    For Each VBCdir In self.arrVBCompDirs
        ThisDir = self.CodeDir & VBCdir
        If Not Fso.FolderExists(ThisDir) Then
            GoTo NextVBCdir
        End If
        Set oFolder = Fso.GetFolder(ThisDir)
        If oFolder.Files.Count = 0 Then
            GoTo NextVBCdir
        End If
        Set oFiles = oFolder.Files
        
        ' Loop over viles
        For Each oFile In oFiles
            If Fso.GetExtensionName(oFile.Name) = "frx" Then
                GoTo NextFile
            End If
            FileNm = Left(oFile.Name, Len(oFile.Name) - 4)
            
            If self.dCurrentVBComps.Exists(FileNm) Then
                Call ImportComponent(Fso, oFile, FileNm, reWsModule)
            Else
                If CheckExisting Then
                    ' Add to dictionary of modules not found among existing
                    NewVBComps.Add oFile.Path, Nothing
                Else
                    ' Simply import
                    self.VBComps.Import oFile.Path
                End If
            End If
NextFile:
        Next oFile
NextVBCdir:
    Next VBCdir
    
    Call CodeDeleteEmptyLines
    
    With NewVBComps
        If .Count > 0 Then
            Debug.Print "Components not found among existing: " & .Count & vbCr _
                & "Add the below components manually:"
            For i = 0 To .Count - 1
                Debug.Print .Keys(i)
            Next i
        End If
    End With
    
    If ShowMsbox Or (Not ShowMsbox And NewVBComps.Count > 0) Then
        MsgBox "Code has been IMPORTED from Git folder." & vbCr & vbCr _
            & "Check unknown components in 'Immediate window': " & NewVBComps.Count, _
            vbInformation, MsbTitle
    End If
    
End Sub

Sub CodeDelete()
    
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    
    For Each VBComp In self.VBComps
        If VBComp.Type = vbext_ct_Document Then
            Set CodeMod = VBComp.CodeModule
            With CodeMod
                If .CountOfLines > 0 Then
                    Call .DeleteLines(1, .CountOfLines)
                End If
            End With
        ElseIf VBComp.Name <> "KeepIt" Then
            self.VBComps.Remove VBComp
        End If
    Next VBComp
    
End Sub

Private Sub CodeDeleteEmptyLines()
' Delete first and last empty lines in code modules.
    
    Dim CodeMod As VBIDE.CodeModule
    Dim VBComp As VBIDE.VBComponent
    
    For Each VBComp In self.VBComps
        Set CodeMod = VBComp.CodeModule
        With CodeMod
            If .CountOfLines > 0 And Len(.Lines(1, 1)) = 0 Then
                ' Delete first empty line.
                Call .DeleteLines(1, 1)
            End If
            If .CountOfLines > 0 Then
                ' Delete empty lines at the end.
                Do While Len(.Lines(.CountOfLines, 1)) = 0
                    Call .DeleteLines(.CountOfLines, 1)
                Loop
            End If
        End With
    Next VBComp
    
End Sub

Private Sub ImportComponent(ByRef Fso As Scripting.FileSystemObject, _
        ByRef oFile As Scripting.File, _
        ByVal FileNm As String, _
        ByRef reWsModule As RegExp)
    
    Dim DelString As String
    Dim FileContents As String
    Dim ReadFile As Object
    Dim VBComp As VBIDE.VBComponent
    Dim Match As Object
    Dim Matches As Object
    
    Set VBComp = self.VBComps.Item(FileNm)
    If VBComp.Type = vbext_ct_Document Then
        ' vbext_ct_Document
        ' For worksheets and ThisWorkbook do this:
        ' remove existing text in code module,
        ' replace with the new one.
        With VBComp.CodeModule
            Set ReadFile = Fso.OpenTextFile(oFile.Path, ForReading)
            FileContents = ReadFile.ReadAll
            
            ' Find correct lines to insert
            Set Matches = reWsModule.Execute(FileContents)
            For Each Match In Matches
                DelString = Match.Value
            Next Match
            FileContents = Replace(FileContents, DelString, vbNullString)
            Call .DeleteLines(1, .CountOfLines)
            Call .InsertLines(1, FileContents)
        End With
    Else
        ' vbext_ct_StdModule, vbext_ct_ClassModule, vbext_ct_MSForm
        ' For Modules, Classes, and Forms do this:
        ' remove existing component, import the new one.
        With self.VBComps
            .Remove VBComp
            .Import oFile.Path
        End With
    End If
    
End Sub

Sub AppClose()
    
    If Not IsWbOpen(self.AppFile) Then
        GoTo Exit_AppClose
    End If
    Workbooks(self.AppFile).Close
    
Exit_AppClose:
    self.IsOpen = IsWbOpen(self.AppFile)
    RbxUI_VDT.Invalidate
    
End Sub

Sub CodeDump()
' Dump code to one text file .vb in CodeDir.
    
    Dim AllCode As String
    Dim FPath As String
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    Dim Fso As New Scripting.FileSystemObject
    Dim oFile As Object
    
    For Each VBComp In self.VBComps
        Set CodeMod = VBComp.CodeModule
        If CodeMod.CountOfLines > 1 Then
            AllCode = AllCode & "' >> " & VBComp.Name _
                & vbCr & vbCr _
                & CodeMod.Lines(1, CodeMod.CountOfLines) _
                & vbCr & vbCr
        End If
    Next VBComp
    
    FPath = self.CodeDir & "\code_dump_" & Format(Now, "yymmdd-hhmmss") & ".vb"
    Set oFile = Fso.CreateTextFile(FPath, ForWriting)
    oFile.WriteLine AllCode
    
    Shell """C:\Program Files\Notepad++\notepad++.exe"" """ & FPath & """", vbNormalFocus
    
End Sub

Sub CodeExport(Optional ByVal IncludeForms As Boolean = False)
    
    Dim FormsCount As Long
    Dim ExportDir As String
    Dim ExportFName As String
    Dim VBComp As VBIDE.VBComponent
    Dim CodeMod As VBIDE.CodeModule
    
    For Each VBComp In self.VBComps
        Select Case VBComp.Type
            Case vbext_ct_ClassModule
                ' classes
                ExportDir = self.CodeDir & dirClassModules
                ExportFName = VBComp.Name & ".cls"
            Case vbext_ct_MSForm
                ' forms
                FormsCount = FormsCount + 1
                If IncludeForms Then
                    ExportDir = self.CodeDir & dirForms
                    ExportFName = VBComp.Name & ".frm"
                Else
                    ' Export forms manually.
                    GoTo NextVBComp
                End If
            Case vbext_ct_StdModule
                ' modules
                If VBComp.Name = "KeepIt" Then
                    GoTo NextVBComp
                End If
                ExportDir = self.CodeDir & dirModules
                ExportFName = VBComp.Name & ".bas"
            Case vbext_ct_Document
                ' sheets
                ExportDir = self.CodeDir & dirObjects
                ExportFName = VBComp.Name & ".cls"
        End Select
        
        Set CodeMod = VBComp.CodeModule
        If CodeMod.CountOfLines > 1 _
                And Not VBComp.Saved Then
            ' Make dir if doesn't exist
            Call MkDirTree(ExportDir)
            ' Only export non-empty modules
            Call VBComp.Export(ExportDir & "\" & ExportFName)
        End If
NextVBComp:
    Next VBComp
    
    Call RemoveEmptyVBCompDirs
    
    MsgBox "The code has been successfully EXPORTED to CodeDir." _
        & vbCr & vbCr _
        & "Workbook: " & self.AppFile & vbCr & vbCr _
        & "Path: " & self.CodeDir & "\<VB components>", _
        vbInformation, MsbTitle
    
    If Not IncludeForms And FormsCount > 0 Then
        MsgBox "You chose to export code WITHOUT USER FORMS." _
            & vbCr & vbCr _
            & "Remember to MANUALLY export them." _
            & vbCr & vbCr _
            & "Forms detected: " & FormsCount & ".", _
            vbExclamation, MsbTitle
    End If
    
    Call RemoveOrphaned
    
End Sub

Private Sub RemoveOrphaned()
    
    Dim OrphanList As String
    Dim Ans As Long
    Dim dKey As Variant
    Dim dOrphaned As New Dictionary
    
    Call GetFilesInCodeDir
    
    ' Check for orphaned files in CodeDir
    For Each dKey In self.dFilesInCodeDir.Keys
        If Not self.dCurrentVBComps.Exists(dKey) Then
            dOrphaned.Add dKey, self.dFilesInCodeDir(dKey)
        End If
    Next dKey
    
    If dOrphaned.Count > 0 Then
        Debug.Print "CodeDir has orhpaned files, check below:"
        For Each dKey In dOrphaned
            Debug.Print dOrphaned(dKey)
            OrphanList = OrphanList & dKey & ", "
        Next dKey
        OrphanList = Left(OrphanList, Len(OrphanList) - 2) & "."
        
        MsgBox "CodeDir contains " & dOrphaned.Count & " orphaned file(s)." _
            & vbCr & vbCr _
            & OrphanList & vbCr & vbCr _
            & "Check their file paths in 'Immediate window':" & vbCr _
            & "Alt+F11 -> Ctrl+G." & vbCr & vbCr _
            & "Please remove them manually, so they are not pushed to Git by mistake!", _
            vbExclamation, MsbTitle
    End If
    
End Sub

Private Sub GetFilesInCodeDir()
    
    Dim FileNm As String
    Dim ExpSubDir As Variant
    Dim Fso As New Scripting.FileSystemObject
    Dim oDir As Scripting.Folder
    Dim oFile As Scripting.File
    Dim oFiles As Scripting.Files
    
    ' Make dictionary of files in CodeDir
    For Each ExpSubDir In self.arrVBCompDirs
        If Fso.FolderExists(self.CodeDir & ExpSubDir) Then
            Set oDir = Fso.GetFolder(self.CodeDir & ExpSubDir)
            Set oFiles = oDir.Files
            For Each oFile In oFiles
                FileNm = oFile.Name
                FileNm = Replace(FileNm, _
                    "." & Fso.GetExtensionName(oFile.Path), vbNullString)
                With self.dFilesInCodeDir
                    If Not .Exists(FileNm) Then
                        .Add FileNm, oFile.Path
                    End If
                End With
            Next oFile
        End If
    Next ExpSubDir
    
End Sub

Private Sub RemoveEmptyVBCompDirs()
' Clean up after export. Remove empty folders.
    
    Dim Fso As New Scripting.FileSystemObject
    Dim oFolder As Scripting.Folder
    Dim IterDir As String
    Dim VBCdir As Variant
    
    ' Loop over component directories and remove with 0 files
    For Each VBCdir In self.arrVBCompDirs
        IterDir = self.CodeDir & VBCdir
        If Fso.FolderExists(IterDir) Then
            Set oFolder = Fso.GetFolder(IterDir)
            If oFolder.Files.Count = 0 Then
                oFolder.Delete
            End If
        End If
    Next VBCdir
    
End Sub

Sub FileBackup()
' Send application file to backup directory.
' Add the existing as a new backup to "older-versions" directory.
    
    Dim OlderVerTarget As String
    Dim OlderVerDir As String
    
    ' Copy postfixed older version to "older-versions"
    Call CopyFileFromTo(self.AppBackupFullName, self.BackupDir & dirOlderVer, True)
Debug.Print "--- Backing up: " & self.AppName & " ---" & vbCr
Debug.Print "From: " & self.AppBackupFullName & vbCr & "To: " & self.BackupDir & dirOlderVer & vbCr
    
    ' Copy latest version to BackupDir.
    Call CopyFileFromTo(self.AppFPath, self.BackupDir)
Debug.Print "From: " & self.AppFPath & vbCr & "To: " & self.BackupDir & vbCr
    
    MsgBox "File: " & self.AppName & vbCr & vbCr _
        & "Current version backed up to local storage, " _
        & "see report in 'Immediate window' in VB Editor (Ctrl+G).", _
        vbInformation, MsbTitle & " | Back up file"
    
End Sub

Sub AppBuild(Optional ByVal WithPostfix As Boolean = True)
' Build app with build-number as postfix "-YYYYMMDD-HHMM" by default.
' Set WithPostfix to False to build with clean name.
' Save to Builds dir.
    
    Dim BuildNo As String
    Dim BuildFullName As String
    Dim BuildAppName As String
    Dim FileExt As String
    Dim Answer As Integer
    Dim FNameExt As Variant
    
    If self.IsOpen Then
        Err.Raise 850, "Build app", "Please close the app before building."
    End If
    Answer = MsgBox("Open bash and check the branch you're on!" _
        & vbCr & vbCr _
        & "Are you sure you want to build from this branch?", _
        vbQuestion + vbOKCancel + vbDefaultButton2, MsbTitle)
    If Answer <> 1 Then
        Exit Sub
    End If
    
    ' Open app, import code.
    Call AppOpenForEditing(ForBuild:=self.IsAppForBuild)
    Call SetVBComponents
    
    If self.AppPwd <> "" Then
        ' Password protect VB Project, use DoEvents!
        LockVBProject(WorkbookName:=self.wb.Name, password:=self.AppPwd) = True
        DoEvents
    End If
    
    FNameExt = Split(self.AppDeployFName, ".", -1, vbTextCompare)
    BuildAppName = FNameExt(0)
    FileExt = "." & FNameExt(1)
    BuildNo = Format(Now, "yyyymmdd-hhmm")
    Call DoBuildNoInjection(BuildNo)
    Call MkDirTree(self.CodeDir & dirBuilds)
    If WithPostfix Then
        BuildFullName = self.CodeDir & dirBuilds & "\" & BuildAppName & "-" & BuildNo & FileExt
    Else
        BuildFullName = self.CodeDir & dirBuilds & "\" & BuildAppName & FileExt
    End If
    Application.EnableEvents = False ' because SaveAs is an event
    Select Case FileExt
        ' Read more on file formats: https://learn.microsoft.com/en-us/office/vba/api/excel.xlfileformat
        Case ".xlam"
            self.wb.SaveAs Filename:=BuildFullName, FileFormat:=xlOpenXMLAddIn
        Case ".xlsb"
            self.wb.SaveAs Filename:=BuildFullName, FileFormat:=xlExcel12
        Case ".xlsm"
            self.wb.SaveAs Filename:=BuildFullName, FileFormat:=xlOpenXMLWorkbookMacroEnabled
        Case Else
            Err.Raise 899, "AppBuild", "Check file extension."
    End Select
    Workbooks(self.wb.Name).Close savechanges:=True
    MsgBox "Build complete:" & vbCr & vbCr & BuildFullName, _
        vbInformation, MsbTitle
    
End Sub

Private Sub DoBuildNoInjection(ByVal BuildNo As String)
' Inject build number into app's hidden sheet.
    
    Dim InjData As Variant
    
    If self.InjectBuild = "" Then
        Exit Sub
    End If
    InjData = Split(self.InjectBuild, ";", -1, vbTextCompare)
    self.wb.Sheets(CInt(InjData(0))).Cells(CInt(InjData(1)), CInt(InjData(2))).Value = BuildNo
    
End Sub

Public Property Get EditXlsm() As Boolean
    EditXlsm = self.EditXlsm
End Property
Public Property Get IsOpen() As Boolean
    IsOpen = self.IsOpen
End Property
Public Property Get AppNm() As String
    AppNm = self.AppName
End Property
Public Property Get AppFPath() As String
    AppFPath = self.AppFPath
End Property
Public Property Let IsOpen(ByVal IsOp As Boolean)
    self.IsOpen = IsOp
End Property
