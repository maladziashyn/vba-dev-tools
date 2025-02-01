Attribute VB_Name = "Dev"
Option Explicit

Sub Rst()
' Check ribbon status
    
    If RbxUI_VDT Is Nothing Then
        Debug.Print "Ribbon is broken."
    Else
        Debug.Print "Ribbon is OK."
    End If
    
End Sub
