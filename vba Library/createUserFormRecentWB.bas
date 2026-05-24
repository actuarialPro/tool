Attribute VB_Name = "createUserFormRecentWB"
Sub CreateWorkbookSelectorForm()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim frm As Object
    Dim ctrl As Object
    
    ' 1. Create the UserForm component
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    
    ' Name the form and set its default dimensions
    'vbComp.name = "RecentWB"
    Set frm = vbComp.Designer
    
    ' Set form dimensions based on your controls' layout boundary
    vbComp.Properties("Caption") = "Workbook Selector"
    vbComp.Properties("Width") = 890
    vbComp.Properties("Height") = 300
    
    ' 2. Add Controls dynamically based on the dataset
    
    ' lblStatus (Label)
    Set ctrl = frm.Controls.Add("Forms.Label.1", "lblStatus")
    With ctrl
        .Caption = "Filter"
        .Left = 12
        .Top = 18
        .Width = 120
        .Height = 18
    End With
    
    ' txtFilter (TextBox)
    Set ctrl = frm.Controls.Add("Forms.TextBox.1", "txtFilter")
    With ctrl
        .Left = 12
        .Top = 48
        .Width = 786
        .Height = 18
    End With
    
    ' lstFiles (ListBox)
    Set ctrl = frm.Controls.Add("Forms.ListBox.1", "lstFiles")
    With ctrl
        .Left = 12
        .Top = 78
        .Width = 788
        .Height = 183.3
    End With
    
    ' cmdOpen (CommandButton)
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "cmdOpen")
    With ctrl
        .Caption = "Open"
        .Left = 810
        .Top = 24
        .Width = 60
        .Height = 24
    End With
    
    ' cmdOpenFolder (CommandButton)
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "cmdOpenFolder")
    With ctrl
        .Caption = "open folder"
        .Left = 810
        .Top = 54
        .Width = 60
        .Height = 24
    End With
    
    ' cmdCancel (CommandButton)
    Set ctrl = frm.Controls.Add("Forms.CommandButton.1", "cmdCancel")
    With ctrl
        .Caption = "Cancel"
        .Left = 810
        .Top = 84
        .Width = 60
        .Height = 24
    End With

    MsgBox "UserForm 'frmWorkbookSelector' created successfully!", vbInformation
End Sub
