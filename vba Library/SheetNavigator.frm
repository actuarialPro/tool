Attribute VB_Name = "SheetNavigator"
Attribute VB_Base = "0{6383F702-D8D1-4274-879A-300AF86212CF}{BB29D18B-EF23-4191-855A-F2E52F3C5972}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
' ==============================================================================
' MODULE-LEVEL DECLARATION
' Place this at the very top of your UserForm code module
' ==============================================================================
Dim mWb As Workbook

' ==============================================================================
' INITIALIZATION & CONFIGURATION EVENTS
' ==============================================================================
Private Sub UserForm_Initialize()
    ' Matches your original intent: act on the currently active workbook.
    ' If you want "the workbook that contains this macro", change to ThisWorkbook.
    Set mWb = ActiveWorkbook
    
    txtFilter.text = vbNullString
    RefreshSheetList
End Sub


' ==============================================================================
' CORE LOGIC & SHEET FILTERING
' ==============================================================================
Private Sub RefreshSheetList()
    Dim ws As Worksheet
    Dim kw As String
    Dim idx As Long
    
    kw = LCase$(Trim$(txtFilter.text))
    lst.Clear
    
    For Each ws In mWb.Worksheets
        ' Requirement: ALWAYS NOT show hidden sheets
        If ws.Visible = xlSheetVisible Then
            If kw = vbNullString Or InStr(1, LCase$(ws.name), kw, vbBinaryCompare) > 0 Then
                idx = lst.ListCount
                lst.AddItem ws.name
                lst.List(idx, 1) = ws.Index ' store sheet index in hidden column
            End If
        End If
    Next ws
    
    ' Auto-select first item so user can type keyword then press Enter
    If lst.ListCount > 0 Then lst.ListIndex = 0
    If lst.ListCount > 0 And kw = "" Then lst.ListIndex = mWb.ActiveSheet.Index - 1
End Sub

Private Sub txtFilter_Change()
    RefreshSheetList
End Sub

' ==============================================================================
' ACTION HANDLERS (SELECTION & NAVIGATION)
' ==============================================================================
Private Sub GoSelectedSheet()
    mWb.Worksheets(lst.Value).Activate
    Unload Me
End Sub

Private Sub lst_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    GoSelectedSheet
End Sub

Private Sub cmdOpen_Click()
    GoSelectedSheet
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

' ==============================================================================
' UX OPTIMIZATION / KEY DOWN INTERCEPTIONS
' ==============================================================================
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' Catch-all placeholder for global form key events if needed
End Sub

Private Sub lst_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Or KeyCode = vbKeySpace Then ' Enter or Space
        KeyCode = 0
        GoSelectedSheet ' Calls selection trigger directly
    End If
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' ? very important: cancel the key so it doesn't insert characters or move cursor
        If lst.ListIndex >= 0 Then
            GoSelectedSheet
        End If
    End If
    
    If KeyCode = vbKeyUp Then
        KeyCode = 0 ' ? very important: cancel the key so it doesn't insert characters or move cursor
        If lst.ListIndex >= 0 Then
            If lst.ListIndex > 0 Then lst.ListIndex = lst.ListIndex - 1
        End If
    End If
    
    If KeyCode = vbKeyDown Then
        KeyCode = 0 ' ? very important: cancel the key so it doesn't insert characters or move cursor
        If lst.ListIndex >= 0 Then
            If lst.ListIndex < lst.ListCount - 1 Then lst.ListIndex = lst.ListIndex + 1
        End If
    End If
End Sub


