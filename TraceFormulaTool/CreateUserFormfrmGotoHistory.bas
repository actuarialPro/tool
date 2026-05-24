Attribute VB_Name = "CreateUserFormfrmGotoHistory"
' Module: modCreateGotoFormUI
Option Explicit

Sub CreateGotoUserFormUI()
    Dim vbComp As VBIDE.VBComponent
    Dim ctrl As Object
    Dim frmName As String
    
    frmName = "frmGotoHistory"
    
    ' Remove existing form if present
    On Error Resume Next
    'ThisWorkbook.VBProject.VBComponents.Remove ThisWorkbook.VBProject.VBComponents(frmName)
    On Error GoTo 0
    
    ' Add a new UserForm
    Set vbComp = ThisWorkbook.VBProject.VBComponents.Add(vbext_ct_MSForm)
    'vbComp.Name = frmName
    
    vbComp.Properties("Width") = 500
    vbComp.Properties("Height") = 300
    With vbComp.Designer.Controls
        ' Label prompt
        Set ctrl = .Add("Forms.Label.1", "lblPrompt", True)
        With ctrl
            .Caption = "Enter address (examples: A1, Sheet1!B2, 'My Sheet'!C3, NamedRange)"
            .Left = 8
            .Top = 8
            .Width = 360
            .Height = 18
        End With
        
        ' TextBox
        Set ctrl = .Add("Forms.TextBox.1", "txtAddress", True)
        With ctrl
            .Left = 8
            .Top = 30
            .Width = 360
            .Height = 20
        End With
        
        ' GoTo button
        Set ctrl = .Add("Forms.CommandButton.1", "btnGoto", True)
        With ctrl
            .Caption = "Go To"
            .Left = 390
            .Top = 28
            .Width = 72
            .Height = 22
        End With
        
        ' ListBox
        Set ctrl = .Add("Forms.ListBox.1", "lstHistory", True)
        With ctrl
            .Left = 8
            .Top = 60
            .Width = 360
            .Height = 140
            .MultiSelect = fmMultiSelectSingle
        End With
        
        ' Back button
        Set ctrl = .Add("Forms.CommandButton.1", "btnBack", True)
        With ctrl
            .Caption = "Back"
            .Left = 8
            .Top = 210
            .Width = 80
            .Height = 24
        End With
        
        ' Clear All button
        Set ctrl = .Add("Forms.CommandButton.1", "btnClearAll", True)
        With ctrl
            .Caption = "Clear All"
            .Left = 96
            .Top = 210
            .Width = 80
            .Height = 24
        End With
        
        ' Status label
        Set ctrl = .Add("Forms.Label.1", "lblStatus", True)
        With ctrl
            .Caption = "Ready"
            .Left = 8
            .Top = 242
            .Width = 360
            .Height = 16
            .ForeColor = &H80000008 ' vbGrayText
        End With
        ' NEW: Add Current Selection button (Upgrade A)
        Set ctrl = .Add("Forms.CommandButton.1", "btnAddCurrent", True)
        With ctrl
            .Caption = "Add Current"
            .Left = 184
            .Top = 210
            .Width = 90
            .Height = 24
        End With
        Set ctrl = .Add("Forms.CommandButton.1", "btnSimplifyFormula", True)
        With ctrl
            .Caption = "Simplify Formula"
            .Left = 390
            .Top = 60
            .Width = 90
            .Height = 24
        End With
    End With
    
    MsgBox "UserForm '" & frmName & "' created with controls. Paste the UserForm code into its module.", vbInformation
End Sub

