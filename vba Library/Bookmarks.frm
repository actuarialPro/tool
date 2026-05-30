Attribute VB_Name = "Bookmarks"
Attribute VB_Base = "0{23EEB12C-25EE-4083-BB91-3CD885C1B657}{B624E82E-6DA6-4A5A-ABF3-E326B6ED3C43}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub lblInstruction_Click()

End Sub

Private Sub UserForm_Initialize()
    InitializeBookmarkSystem
    UI_RefreshList
    If BookmarkGetMode Then
        btnGoTo.Caption = "Go To (Enter)"
        CreateShortcutButtons
    Else
        btnSet.Caption = "Set Bookmark (Enter)"
    End If
End Sub

Private Sub lstBookmarks_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If lstBookmarks.ListIndex = -1 Then Exit Sub
    
    If KeyCode = vbKeyReturn Then   ' Enter
        KeyCode = 0
        If BookmarkGetMode Then
            btnGoTo_Click
        Else
            btnSet_Click
        End If
        
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
    
    
' 3. Handle Shift + Space (Go to Row only)
    If KeyCode = vbKeySpace And BookmarkGetMode And Shift = 1 Then
        KeyCode = 0
        GoToCell RowOnly:=True, ColOnly:=False
        Unload Me
        Exit Sub
    End If

    ' 4. Handle Ctrl + Space (Go to Column only)
    If KeyCode = vbKeySpace And BookmarkGetMode And Shift = 2 Then
        KeyCode = 0
        GoToCell RowOnly:=False, ColOnly:=True
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub UI_RefreshList()
    Dim i As Long
    Dim displayStr As String
    Dim cellPreviewText As String
    Dim idx As Long
    Dim targetWs As Worksheet
    
    idx = lstBookmarks.ListIndex
    lstBookmarks.Clear
    
    For i = 1 To 9
        cellPreviewText = "" ' Reset buffer for every slot step
        
        If BookmarkHistory(i).SheetName = "" Then
            displayStr = i & ": (empty)"
        ElseIf BookmarkHistory(i).WorkbookName <> ActiveWorkbook.name Then
            displayStr = i & ": [used by other workbook]"
        Else
            displayStr = i & ": " & BookmarkHistory(i).SheetName & "!" & BookmarkHistory(i).Address
            
            ' Fetch cell contents safely contextually
            On Error Resume Next
            Set targetWs = Worksheets(BookmarkHistory(i).SheetName)
            If Not targetWs Is Nothing Then
                ' Grab only the first cell text to accommodate large multi-cell selection blocks gracefully
                cellPreviewText = targetWs.Range(BookmarkHistory(i).Address).Cells(1, 1).text
            End If
            On Error GoTo 0
        End If
        
        ' Add item creates the new row index, defining Column 0 natively
        lstBookmarks.AddItem displayStr
        
        ' Map preview string to Column 1 (the 2nd structural column visual block)
        If cellPreviewText <> "" Then
            lstBookmarks.List(lstBookmarks.ListCount - 1, 1) = "[" & cellPreviewText & "]"
        ElseIf BookmarkHistory(i).SheetName <> "" And BookmarkHistory(i).WorkbookName = ActiveWorkbook.name Then
            lstBookmarks.List(lstBookmarks.ListCount - 1, 1) = "[Blank]"
        End If
    Next i
    
    If idx >= 0 Then lstBookmarks.ListIndex = idx Else lstBookmarks.ListIndex = 0
End Sub

Private Sub btnSet_Click()
    Dim slot As Long
    
    slot = lstBookmarks.ListIndex + 1
    If slot < 1 Or slot > 9 Then Exit Sub
    
    BookmarkHistory(slot).WorkbookName = ActiveWorkbook.name
    BookmarkHistory(slot).SheetName = ActiveSheet.name
    BookmarkHistory(slot).Address = Selection.Address
    
    'UI_RefreshList
    Unload Me
End Sub

Private Sub btnGoTo_Click()
    GoToCell RowOnly:=False, ColOnly:=False
    Unload Me
End Sub

Private Sub GoToCell(RowOnly As Boolean, ColOnly As Boolean)
    Dim slot As Long
    Dim ws As Worksheet
    Dim rng As Range
    
    slot = lstBookmarks.ListIndex + 1
    If slot < 1 Or slot > 9 Then Exit Sub
    
    If BookmarkHistory(slot).SheetName = "" Then
        MsgBox "Slot is empty.", vbExclamation, "Alert"
        Exit Sub
    ElseIf BookmarkHistory(slot).WorkbookName <> ActiveWorkbook.name Then
        MsgBox "Bookmark belongs to another active file context.", vbExclamation, "Forbidden"
        Exit Sub
    End If
    
    On Error Resume Next
    Set ws = Worksheets(BookmarkHistory(slot).SheetName)
    Set rng = ws.Range(BookmarkHistory(slot).Address)
    On Error GoTo 0
    
    ' Ensure the referenced sheet and range still exist safely
    If ws Is Nothing Or rng Is Nothing Then
        MsgBox "Worksheet/Address range reference missing or altered"
        Exit Sub
    End If
    
    ' Standard Navigation: Switch tabs and move to full address
    If Not RowOnly And Not ColOnly Then
        ws.Activate
        rng.Select
        Application.Goto rng
        Exit Sub
    End If
    
    
    ' Shift + Space: Row only (Keep current active column)
    If RowOnly Then
        ActiveSheet.Cells(rng.Row, ActiveCell.Column).Select
    ' Ctrl + Space: Column only (Keep current active row)
    ElseIf ColOnly Then
        ActiveSheet.Cells(ActiveCell.Row, rng.Column).Select
    End If
End Sub



' Helper to dynamically generate the new shortcut buttons in Get Mode
Private Sub CreateShortcutButtons()
    Dim btnShiftSpace As MSForms.CommandButton
    Dim btnCtrlSpace As MSForms.CommandButton
    Dim btnWidth As Double, btnHeight As Double, btnTop As Double
    
    ' Match dimensions and horizontal alignment of your existing btnGoTo
    btnWidth = btnGoTo.Width
    btnHeight = btnGoTo.Height
    btnTop = btnGoTo.Top + btnHeight + 5 ' Position them neatly below the GoTo button
    
    ' 1. Create Shift + Space Button (Row Only)
    Set btnShiftSpace = Me.Controls.Add("Forms.CommandButton.1", "btnShiftSpace", True)
    With btnShiftSpace
        .Caption = "GoTo Row (Shift+Space)"
        .Left = btnGoTo.Left
        .Top = btnTop
        .Width = btnWidth ' Split the width across two equal halves
        .Height = btnHeight
        .TakeFocusOnClick = False
    End With
    
    ' 2. Create Ctrl + Space Button (Col Only)
    Set btnCtrlSpace = Me.Controls.Add("Forms.CommandButton.1", "btnCtrlSpace", True)
    With btnCtrlSpace
        .Caption = "GoTo Col (Ctrl+Space)"
        .Left = btnGoTo.Left
        .Top = btnTop + btnHeight + 5
        .Width = btnWidth
        .Height = btnHeight
        .TakeFocusOnClick = False
    End With
    
    ' Automatically expand the form height slightly to fit the new buttons safely
    Me.Height = Me.Height + 2 * btnHeight + 10
End Sub

' Catch clicks on the dynamically generated buttons natively via Control_Click
Private Sub Me_OnControlClick(ByVal ControlName As String)
    If lstBookmarks.ListIndex = -1 Then Exit Sub
    
    If ControlName = "btnShiftSpace" Then
        GoToCell RowOnly:=True, ColOnly:=False
        Unload Me
    ElseIf ControlName = "btnCtrlSpace" Then
        GoToCell RowOnly:=False, ColOnly:=True
        Unload Me
    End If
End Sub

