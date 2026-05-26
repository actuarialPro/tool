Attribute VB_Name = "Bookmarks"
Attribute VB_Base = "0{55B75D86-F06F-4580-A538-05FFF3AEB40A}{CA077123-3D01-4628-9B62-9431F0D33AF5}"
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
    Else
        btnSet.Caption = "Set Bookmark (Enter)"
    End If
End Sub

Private Sub lstBookmarks_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then   ' Enter
        
        KeyCode = 0
        If BookmarkGetMode Then
            btnGoTo_Click
        Else
            btnSet_Click
        End If
        lstBookmarks.SetFocus
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
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
    
    UI_RefreshList
End Sub

Private Sub btnGoTo_Click()
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
    On Error GoTo 0
    
    If ws Is Nothing Then
        MsgBox "Sheet tracking reference lost.", vbCritical, "Error"
        Exit Sub
    End If
    
    On Error Resume Next
    ws.Activate
    Set rng = ws.Range(BookmarkHistory(slot).Address)
    On Error GoTo 0
    
    If Not rng Is Nothing Then
        rng.Select
        Application.Goto rng
    Else
        MsgBox "Address range reference missing or altered.", vbCritical, "Invalid Range"
    End If
End Sub

