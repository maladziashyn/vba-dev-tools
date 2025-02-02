Attribute VB_Name = "ButtonActions"
Option Explicit

Sub DoCodeFileAction(ByVal ActionType As String)
    
    On Error GoTo Err_DoCodeFileAction
    Call OnStart
    Set oApp = New App
    
    Select Case ActionType
        Case "open"
            Call oApp.AppOpenForEditing
        Case "open xlsm"
            Workbooks.Open oApp.AppFPath
            oApp.IsOpen = True
            RbxUI_VDT.Invalidate
        Case "import"
            Call oApp.CodeImport
        Case "delete"
            Call oApp.CodeDelete
        Case "dump"
            Call oApp.CodeDump
        Case "export exclude forms"
            Call oApp.CodeExport
        Case "export include forms"
            Call oApp.CodeExport(IncludeForms:=True)
        Case "backup"
            Call oApp.FileBackup
        Case "build postfix"
            Call oApp.AppBuild
        Case "build no postfix"
            Call oApp.AppBuild(WithPostfix:=False)
        
        Case "close"
            Call oApp.AppClose
        
    End Select
    
Err_DoCodeFileAction:
    Call OnExit
    
End Sub
