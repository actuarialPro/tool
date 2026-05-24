Attribute VB_Name = "createUserFormBookmarks"
Sub BuildMultiColumnBookmarkSystem()
    Dim vbProj As Object
    Dim vbCompForm As Object
    Dim newForm As Object
    Dim lblTitle As Object, lstBox As Object, btnSet As Object, btnGoTo As Object
    
    Set vbProj = ThisWorkbook.VBProject
    
    ' Safely delete old form if it exists to refresh design properties cleanly
    On Error Resume Next
    On Error GoTo 0
    
    ' 1. Create the expanded UserForm container
    Set vbCompForm = vbProj.VBComponents.Add(3) ' 3 = vbext_ct_MSForm
    With vbCompForm
        '.Name = "Bookmarks"
        .Properties("Caption") = "Multi-Column Bookmark Manager"
        .Properties("Width") = 360 ' Widened to make room for cell preview text
        .Properties("Height") = 240
    End With
    
    Set newForm = vbCompForm.Designer
    
    ' 2. Add Header Label
    Set lblTitle = newForm.Controls.Add("Forms.Label.1")
    With lblTitle
        .name = "lblInstruction"
        .Caption = "Slot Location / Reference:"
        .Left = 12: .Top = 10: .Width = 330: .Height = 15
    End With
    
    ' 3. Add Multi-Column ListBox
    Set lstBox = newForm.Controls.Add("Forms.ListBox.1")
    With lstBox
        .name = "lstBookmarks"
        .Left = 12: .Top = 28: .Width = 320: .Height = 135
        .ColumnCount = 2 ' Crucial property change
        .ColumnWidths = "200 pt;110 pt" ' Allocating specific spatial splits
    End With
    
    ' 4. Add Control Buttons
    Set btnSet = newForm.Controls.Add("Forms.CommandButton.1")
    With btnSet
        .name = "btnSet": .Caption = "Set Bookmark": .Left = 12: .Top = 175: .Width = 145: .Height = 26
    End With
    
    Set btnGoTo = newForm.Controls.Add("Forms.CommandButton.1")
    With btnGoTo
        .name = "btnGoTo": .Caption = "Go To": .Left = 187: .Top = 175: .Width = 145: .Height = 26
    End With
    
    MsgBox "Multi-column userform layout compiled!", vbInformation
End Sub

