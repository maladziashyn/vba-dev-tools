Attribute VB_Name = "Ribbon"
Option Explicit

Public RbxUI_VDT As IRibbonUI

Sub VDT_OnLoad(ByRef ribbon As IRibbonUI)
' Load custom ribbon tab.
    
    SelItmIdx = 0 ' default app index
    SelectedApp = GetSelectedApp
    Set oApp = New App
    
    Set RbxUI_VDT = ribbon
    RbxUI_VDT.ActivateTab "tabVDT"
    Application.WindowState = xlMaximized
    
End Sub

Sub VDT_ClickButton(ByRef control As IRibbonControl)
    
    Select Case control.id
        Case "btnExit"
            ThisWorkbook.Close savechanges:=True
        Case "btnOpenEdit"
            Call DoCodeFileAction("open")
        Case "btnCloseApp"
            Call DoCodeFileAction("close")
        Case "btnImportCode"
            Call DoCodeFileAction("import")
        Case "btnDelCode"
            Call DoCodeFileAction("delete")
        Case "btnDumpCode"
            Call DoCodeFileAction("dump")
        Case "btnExportCode", "itemExportExclForms"
            Call DoCodeFileAction("export exclude forms")
        Case "itemExportInclForms"
            Call DoCodeFileAction("export include forms")
    End Select
    
End Sub

Sub VDT_ClickButton_WithGetPressed(ByRef control As IRibbonControl, ByRef pressed As Boolean)
' Turn AddIn mode on/off.
' Turn it off to make changes on wsMain.
    
    With ThisWorkbook
        If .IsAddin Then
            .IsAddin = False
        Else
            .IsAddin = True
            .Save
        End If
        pressed = .IsAddin
    End With
        
    ' Refresh ribbon
    If RbxUI_VDT Is Nothing Then
        MsgBox "Error: Custom ribbon tab was reset because of an error. " _
            & "Restart VBA DevTools to restore the tab.", vbCritical, MsbTitle
    Else
        RbxUI_VDT.Invalidate
    End If
    
End Sub

Sub VDT_GetPressed(ByRef control As IRibbonControl, ByRef returnedVal)
    
    Select Case control.Tag
        Case "AddInMode"
            returnedVal = ThisWorkbook.IsAddin
    End Select
    
End Sub

Sub VDT_GetEnabled(ByRef control As IRibbonControl, ByRef returnedVal)
' Enable/disable buttons depending on AddIn mode on/off.
    
    If ThisWorkbook.IsAddin Then
        Select Case control.Tag
            Case "gpCode"
                returnedVal = oApp.IsOpen And oApp.AppNm <> "VBADevTools"
            Case "gpCodeDump", "gpCodeExport"
                returnedVal = oApp.IsOpen
            Case "gpOpen"
                returnedVal = Not oApp.IsOpen
            Case "gpOpenXlsm"
                returnedVal = Not oApp.IsOpen And oApp.EditXlsm
            Case "gpClose"
                returnedVal = oApp.IsOpen And oApp.AppNm <> "VBADevTools"
            Case Else
                returnedVal = True
        End Select
    Else
        returnedVal = False
    End If
    
End Sub

Sub VDT_DropDown_GetItemCount(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = wsMain.ListObjects("tblApps").DataBodyRange.Rows.Count
    
End Sub

Sub VDT_DropDown_GetItemLabel(ByRef control As IRibbonControl, ByRef index As Integer, ByRef returnedVal)
    
    returnedVal = wsMain.ListObjects("tblApps").DataBodyRange(index + 1, 1)
    
End Sub

Sub VDT_DropDown_GetSelectedItemIndex(ByRef control As IRibbonControl, ByRef returnedVal)
    
    returnedVal = SelItmIdx
    
End Sub

Sub VDT_DropDown_OnAction(ByRef control As IRibbonControl, ByRef id As String, ByRef index As Integer)
' Store selected app in Config type.
    
    SelItmIdx = index
    SelectedApp = GetSelectedApp
    Set oApp = New App
    RbxUI_VDT.Invalidate
    
End Sub

Private Function GetSelectedApp() As String
    
    GetSelectedApp = wsMain.ListObjects("tblApps").DataBodyRange(SelItmIdx + 1, 1).Value
    
End Function
