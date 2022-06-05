Attribute VB_Name = "mod99_Ribbon"
Option Explicit

'Public g_RibbonUI As IRibbonUI

' Ribbon buttons
Public Sub RbnAddWhiteBlock(control As IRibbonControl)
    AddWhiteBlock
End Sub

Public Sub RbnAddSemiBlock(control As IRibbonControl)
    AddSemiBlock
End Sub

' Coding to - someday - disable buttons when they can't be used
'Public Sub RibbonLoaded(ribbon As IRibbonUI)
'    Set g_RibbonUI = ribbon
'End Sub

'Function ButtonGetEnabled(control As IRibbonControl, ByRef returnedVal)
    
    ' Disable buttons if user isn't in appropriate mode
    'If Application.Windows.Count = 0 Or (ActiveWindow.ViewType <> ppViewNormal And ActiveWindow.ViewType <> ppViewSlide) Then
    '    returnedVal = False
    'ElseIf ActiveWindow.Selection.SlideRange.Count > 2 Then
    '    returnedVal = False
    'Else
    '    returnedVal = True
    'End If

'End Function

'Private Sub AppEvents_WindowSelectionChange(ByVal Sel As Selection)
    'Call g_RibbonUI.InvalidateControl("btnAddWhiteBlock")
    'Call g_RibbonUI.InvalidateControl("btnAddSemiTransBlock")
'End Sub


