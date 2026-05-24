Attribute VB_Name = "createUserFormWorkbookSelector"
Sub CreateWorkbookSelectorForm()
    Dim vbProj As Object
    Dim vbComp As Object
    Dim formCodeModule As Object
    Dim ctrl As Object
    Dim lineNum As Long
    
    ' 1. Create the UserForm component
    Set vbProj = ThisWorkbook.VBProject
    Set vbComp = vbProj.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    'vbComp.name = "frmWorkbookSelector"
    
    ' 2. Set form properties
    With vbComp.Properties
        .item("Caption") = "Goto Workbook"
        .item("Width") = 580
        .item("Height") = 320
    End With
    
    ' 3. Add Title Label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .name = "lblTitle"
        .Caption = "Open Workbooks:"
        .Left = 12
        .Top = 12
        .Width = 440
        .Height = 15
        .Font.Bold = True
    End With
    
    ' 4. Add Filter Label
    Set ctrl = vbComp.Designer.Controls.Add("Forms.Label.1")
    With ctrl
        .name = "lblFilter"
        .Caption = "Filter:"
        .Left = 12
        .Top = 35
        .Width = 440
        .Height = 15
    End With
    
    ' 5. Add Filter TextBox
    Set ctrl = vbComp.Designer.Controls.Add("Forms.TextBox.1")
    With ctrl
        .name = "txtFilter"
        .Left = 12
        .Top = 52
        .Width = 440
        .Height = 20
    End With
    
    ' 6. Add ListBox
    Set ctrl = vbComp.Designer.Controls.Add("Forms.ListBox.1")
    With ctrl
        .name = "lstworkbooks"
        .Left = 12
        .Top = 80
        .Width = 440
        .Height = 200
    End With
    
    ' 7. Add OK Button
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .name = "cmdOK"
        .Caption = "OK"
        .Left = 470
        .Top = 40
        .Width = 80
        .Height = 25
        .Default = True
    End With
    
    ' 8. Add Cancel Button
    Set ctrl = vbComp.Designer.Controls.Add("Forms.CommandButton.1")
    With ctrl
        .name = "cmdCancel"
        .Caption = "Cancel"
        .Left = 470
        .Top = 75
        .Width = 80
        .Height = 25
        .Cancel = True
    End With
    

End Sub
