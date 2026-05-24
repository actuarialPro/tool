Attribute VB_Name = "createUserFormSheetNavigator"
Sub CreateSheetNavigatorForm()
    Dim vbProj As Object
    Dim frm As Object
    Dim ctrlTxt As Object, ctrlLst As Object
    Dim ctrlBtnOpen As Object, ctrlBtnCancel As Object
    Dim codeLines As String
    
    ' Set reference to the active workbook's VB Project
    Set vbProj = ThisWorkbook.VBProject
    ' 1. Create the new UserForm (3 = vbext_ct_MSForm)
    Set frm = vbProj.VBComponents.Add(3)
    'frm.name = "frmSheetNavigator"
    ' Configure the Form Properties
    With frm
        .Properties("Caption") = "Goto Sheet"
        .Properties("Width") = 670
        .Properties("Height") = 470
    End With
    ' 1.2. Add Title Label
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .name = "lblTitle"
        .Caption = "Open Worksheets:"
        .Left = 12
        .Top = 12
        .Width = 440
        .Height = 15
        .Font.Bold = True
    End With
    
    ' 1.3. Add Filter Label
    Set ctrl = frm.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .name = "lblFilter"
        .Caption = "Filter:"
        .Left = 12
        .Top = 35
        .Width = 440
        .Height = 15
    End With
    ' 2. Add and Configure txtFilter (TextBox)
    Set ctrlTxt = frm.Designer.Controls.Add("Forms.TextBox.1")
    With ctrlTxt
        .name = "txtFilter"
        .Left = 12
        .Top = 66
        .Width = 516
        .Height = 20
    End With
    
    ' 3. Add and Configure lst (ListBox)
    Set ctrlLst = frm.Designer.Controls.Add("Forms.ListBox.1")
    With ctrlLst
        .name = "lst"
        .Left = 12
        .Top = 96
        .Width = 516
        .Height = 330
    End With
    
    ' 4. Add and Configure cmdOpen (CommandButton)
    Set ctrlBtnOpen = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrlBtnOpen
        .name = "cmdOpen"
        .Caption = "Open"
        .Left = 552
        .Top = 66
        .Width = 80
        .Height = 24
        .Default = True ' Allows pressing Enter key globally to trigger it
    End With
    
    ' 5. Add and Configure cmdCancel (CommandButton)
    Set ctrlBtnCancel = frm.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrlBtnCancel
        .name = "cmdCancel"
        .Caption = "Cancel"
        .Left = 552
        .Top = 102
        .Width = 80
        .Height = 24
        .Cancel = True ' Allows pressing Esc key globally to close form
    End With
End Sub
