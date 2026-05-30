Attribute VB_Name = "frmPrecedencyTracer"
Attribute VB_Base = "0{F438EB52-8210-4968-8B9D-BEEAC0D66142}{C8596208-1575-4026-89C9-C7AD171B9454}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False

' --- GLOBAL STATE FOR THE USERFORM SESSION ---
Private g_ActiveTarget As Range     ' Cell currently being analyzed

Private Sub lstPrecedents_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
End Sub

Private Sub UserForm_Initialize()
    ' Configure ListBox for 2 columns (Address and Value)
    lstPrecedents.ColumnCount = 2
    lstPrecedents.ColumnWidths = "275pt;75pt" ' Adjust widths as needed
    ' Automatically target the currently active cell if it is valid
    If TypeOf Selection Is Range Then
        Set g_ActiveTarget = ActiveCell
        UpdateActiveCellView g_ActiveTarget
        lstPrecedents.listIndex = Application.Min(lstPrecedents.ListCount - 1, 0)
    Else
        lblStatus.Caption = "Select an active cell and click 'Trace Current'."
    End If
End Sub

' --- DOUBLE CLICK PRECEDENT (FAST NAVIGATOR & AUTOMATIC DRILL DOWN) ---
Private Sub lstPrecedents_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIdx As Long
    selectedIdx = lstPrecedents.listIndex
    If selectedIdx = -1 Then Exit Sub

    Dim targetRange As Range
    Set targetRange = ResolvePrecedentFromList(selectedIdx)

    If Not targetRange Is Nothing Then
        ' Navigate focus directly to the target worksheet/cell
        NavigateToRange targetRange
    Else
        lblStatus.Caption = "Error navigating to selected cell."
    End If

    Unload Me
    
    
End Sub

' ==============================================================================
' --- PRIVATE DEPENDENCY CRAWLING ENGINE ---
' ==============================================================================

Private Sub TracePrecedents(targetRange As Range)
    lstPrecedents.Clear

    If targetRange Is Nothing Then Exit Sub

    ' If cell doesn't have formulas, it has no precedents
    If Not targetRange.HasFormula Then
        lblStatus.Caption = "Cell has no formula."
        Exit Sub
    End If

    Dim precedentsColl As Collection
    Dim originalScreenUpdating As Boolean
    
    originalScreenUpdating = Application.ScreenUpdating
    Application.ScreenUpdating = False

    ' Get the precedents using our simplified collection-based helper function
    Set precedentsColl = AllPrecedentsOf(targetRange)
    
    Application.ScreenUpdating = originalScreenUpdating

    ' --- POPULATE THE LISTBOX ---
    If precedentsColl.Count = 0 Then
        lblStatus.Caption = "No direct precedents resolved."
        Exit Sub
    End If

    Dim precedentItem As Variant
    Dim cellRef As Range

    For Each precedentItem In precedentsColl
        Set cellRef = precedentItem
        lstPrecedents.AddItem GetFriendlyAddress(cellRef)
        
        ' 2. Evaluate the cell text in real-time, handling multi-cell scenarios safely
        lstPrecedents.List(lstPrecedents.ListCount - 1, 1) = cellRef.Cells(1, 1).Text
    Next precedentItem

    lblStatus.Caption = "Found " & precedentsColl.Count & " direct precedent(s)."
End Sub

' --- CORE PRECEDENT CRAWLING HELPER FUNCTION ---
' Adapts your simplified dual-path arrow exploration to return a Collection of Range objects.
Private Function AllPrecedentsOf(aCell As Range) As Collection
    Dim precedentsColl As New Collection
    Dim testRange As Range
    Dim i As Long

    Set aCell = aCell.Cells(1, 1)

    ' Explicitly activate parent workbook and sheet to ensure arrow context is correct
    On Error Resume Next
    aCell.Worksheet.Parent.Activate
    aCell.Worksheet.Activate
    aCell.Select
    On Error GoTo 0

    aCell.Worksheet.ClearArrows
    aCell.ShowPrecedents

    ' --- PATH 1: Link-based loop on Arrow 1 (Handles cross-sheet/external reference paths) ---
    i = 0
    Do
        i = i + 1
        
        ' Reset navigation context back to parent range prior to checking step
        On Error Resume Next
        aCell.Worksheet.Parent.Activate
        aCell.Worksheet.Activate
        aCell.Select
        On Error GoTo 0
        
        On Error Resume Next
        aCell.NavigateArrow True, 1, i
        'If ActiveCell.Address(, , , True) = aCell.Address(, , , True) Then Exit Do
        If Err Then Exit Do
        On Error GoTo 0
        
        Set testRange = aCell.NavigateArrow(True, 1, i)
        
        ' Add to collection using unique key to automatically ignore duplicate overlaps
        On Error Resume Next
            If ActiveCell.Address(, , , True) <> aCell.Address(, , , True) Then
                precedentsColl.Add testRange, testRange.Address(External:=True)
            End If
        On Error GoTo 0
    Loop 'Until ActiveCell.Address(, , , True) = aCell.Address(, , , True)

    ' --- PATH 2: Arrow-based loop on Link 1 (Handles local worksheet reference paths) ---
    i = 1
    Do
        i = i + 1
        
        ' Reset navigation context back to parent range prior to checking step
        On Error Resume Next
        aCell.Worksheet.Parent.Activate
        aCell.Worksheet.Activate
        aCell.Select
        On Error GoTo 0
        
        On Error Resume Next
        aCell.NavigateArrow True, i, 1
        If ActiveCell.Address(, , , True) = aCell.Address(, , , True) Then Exit Do
        On Error GoTo 0
        
        Set testRange = aCell.NavigateArrow(True, i, 1)
        
        ' Add to collection using unique key to automatically ignore duplicate overlaps
        On Error Resume Next
        precedentsColl.Add testRange, testRange.Address(External:=True)
        On Error GoTo 0
    Loop Until ActiveCell.Address(, , , True) = aCell.Address(, , , True)

    ' Cleanup blue precedent lines on active sheet before completing
    On Error Resume Next
    aCell.Worksheet.Parent.Activate
    aCell.Worksheet.Activate
    aCell.Select
    aCell.Worksheet.ClearArrows
    On Error GoTo 0

    Set AllPrecedentsOf = precedentsColl
End Function

' --- AUXILIARY VIEWPORT SYNCHRONIZER ---
Private Sub UpdateActiveCellView(targetRange As Range)
    Dim friendlyAddr As String
    friendlyAddr = GetFriendlyAddress(targetRange)

    ' Update the visual labels and the form title dynamically
    lblTargetPrompt.Caption = "Traced Cell: " & friendlyAddr
    Me.Caption = "Tracer: " & friendlyAddr

    ' Parse dependencies
    TracePrecedents targetRange
End Sub

' --- COMPACT NAVIGATION ENGINE ---
Private Sub NavigateToRange(targetRange As Range)
    On Error Resume Next
    ' Ensure Workbook and Sheet are active prior to jumping
    targetRange.Worksheet.Parent.Activate
    targetRange.Worksheet.Activate
    Application.Goto targetRange, Scroll:=False
    On Error GoTo 0
End Sub

' --- RECONSTRUCT RANGE REFERENCE FROM LIST SELECTION ---
Private Function ResolvePrecedentFromList(listIdx As Long) As Range
    On Error Resume Next
    Dim addressStr As String
    addressStr = lstPrecedents.List(listIdx)

    ' Run evaluation of the address string to build a valid Range object context
    Set ResolvePrecedentFromList = Application.Evaluate(addressStr)
    On Error GoTo 0
End Function

' --- ADAPTIVE ADDRESS CONVERTER FOR CROSS-SHEET READABILITY ---
Private Function GetFriendlyAddress(targetRange As Range) As String
    GetFriendlyAddress = targetRange.Worksheet.Name & "!" & targetRange.Address(False, False)
    ' If cross-workbook, prefix with brackets
    If targetRange.Worksheet.Parent.Name <> ActiveWorkbook.Name Then GetFriendlyAddress = "[" & targetRange.Worksheet.Parent.Name & "]" & GetFriendlyAddress
End Function

