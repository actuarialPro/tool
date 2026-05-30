Attribute VB_Name = "CreatePrecedencyTracer"
' ==============================================================================
' Module: modCreatePrecedencyTracerUI
' Description: Programmatically generates an elegant, ultra-compact UserForm
'              interface for tracing cell precedents.
' ==============================================================================
Option Explicit

Sub CreatePrecedencyTracerUI()
    Dim vbComp As VBIDE.VBComponent
    Dim ctrl As Object
    Dim frmName As String

    frmName = "frmPrecedencyTracer"

    ' Remove existing form if present
    On Error Resume Next
    ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(frmName)
    On Error GoTo 0

    ' Add a new UserForm
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    'vbComp.Name = frmName

    ' Design canvas parameters - ultra-compact height profile
    vbComp.Properties("Width") = 400
    vbComp.Properties("Height") = 260
    vbComp.Properties("Caption") = "Cell Precedency Tracer"

    With vbComp.Designer.Controls
        ' Label prompt displaying currently traced cell reference
        Set ctrl = .Add("Forms.Label.1", "lblTargetPrompt", True)
        With ctrl
            .Caption = "Traced Cell: [None]"
            .Left = 8
            .Top = 10
            .Width = 360
            .Height = 14
            .Font.Bold = True
        End With
        
        
        ' Label prompt for Precedents List
        Set ctrl = .Add("Forms.Label.1", "lblPrecedentsPrompt", True)
        With ctrl
            .Caption = "Direct Precedents (Double-click any item to navigate):"
            .Left = 8
            .Top = 34
            .Width = 360
            .Height = 14
        End With
        
        ' ListBox for displaying precedents (Multi-column layout)
        Set ctrl = .Add("Forms.ListBox.1", "lstPrecedents", True)
        With ctrl
            .Left = 8
            .Top = 50
            .Width = 360
            .Height = 150
            .MultiSelect = fmMultiSelectSingle
        End With
        
        
        ' Status label
        Set ctrl = .Add("Forms.Label.1", "lblStatus", True)
        With ctrl
            .Caption = "Ready"
            .Left = 8
            .Top = 206
            .Width = 360
            .Height = 16
            .ForeColor = &H80000008 ' vbGrayText
        End With
    End With

    MsgBox "UserForm '" & frmName & "' created. Paste the UserForm code into its module.", vbInformation
End Sub


