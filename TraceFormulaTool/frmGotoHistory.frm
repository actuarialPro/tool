Attribute VB_Name = "frmGotoHistory"
Attribute VB_Base = "0{AF8AE938-40F0-4FC4-BD53-AF6250152C65}{767A0AF1-975B-4FB9-A7E6-CE455204B186}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit

Private Sub UserForm_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

End Sub

Private Sub UserForm_Initialize()
    ' Initialize global stack on first run if it doesn't exist yet
    If g_HistoryStack Is Nothing Then Set g_HistoryStack = New Collection
    
    ' Reload existing history elements from the Excel session memory
    RefreshListBox
    lblStatus.Caption = "Ready"
End Sub
' --- GLOBAL FORM KEY MONITORING (SHORTCUTS) ---
Private Sub UserForm_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' 1. ESC KEY -> Exit the form cleanly
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    
    ' 2. ALT + LEFT ARROW KEY -> Trigger the Back functionality
    ' (Shift = 4 means the Alt key is being held down)
    If Shift = 4 And KeyCode = vbKeyLeft Then
        KeyCode = 0 ' Consume the key event
        btnBack_Click
        Exit Sub
    End If
End Sub

' --- ENTER KEY HANDLING ---
Private Sub txtAddress_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    ' If user presses Enter inside the TextBox, trigger GoTo
    If KeyCode = vbKeyEscape Then
        Unload Me
        Exit Sub
    End If
    If KeyCode = vbKeyReturn Then
        KeyCode = 0 ' Consume the key
        btnGoto_Click
    End If
End Sub

' --- ADD CURRENT SELECTION BUTTON CLICK (UPGRADE A) ---
Private Sub btnAddCurrent_Click()
    ' Ensure the user's active focus is actually an Excel cell selection range
    If TypeOf Selection Is Range Then
        PushToGlobalStack Selection
        RefreshListBox
        lblStatus.Caption = "Bookmarked: " & GetFriendlyAddress(Selection)
    Else
        lblStatus.Caption = "Selection must be a standard cell range."
    End If
End Sub

' --- GO TO BUTTON CLICK ---
Private Sub btnGoto_Click()
    Dim addressInput As String
    addressInput = Trim(txtAddress.Text)
    
    If addressInput = "" Then
        lblStatus.Caption = "Please enter a valid address or Named Range."
        Exit Sub
    End If
    
    If Left(addressInput, 7) = "VLOOKUP" Then addressInput = evalStr(addressInput)
    ' Resolve the address using Evaluate
    Dim evalResult As Variant
    On Error Resume Next
    Set evalResult = Application.Evaluate(addressInput)
    On Error GoTo 0
    
    
    Dim targetRange As Range
    If Not IsError(evalResult) Then
        If TypeOf evalResult Is Range Then
            Set targetRange = evalResult
        End If
    End If
    
    If targetRange Is Nothing Then
        lblStatus.Caption = "Error: Invalid address or range name."
        Exit Sub
    End If
    
    ' Execute navigation
    On Error Resume Next
    Application.Goto targetRange, Scroll:=True
    If Err.Number <> 0 Then
        lblStatus.Caption = "Error navigating: " & Err.Description
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Push to the global stack and sync UI
    PushToGlobalStack targetRange
    RefreshListBox
    
    txtAddress.Text = ""
    lblStatus.Caption = "Moved to: " & GetFriendlyAddress(targetRange)
End Sub

' --- BACK BUTTON CLICK (UNWIND STACK) ---
Private Sub btnBack_Click()
    Dim selectedIndex As Long
    selectedIndex = lstHistory.listIndex
    
    If selectedIndex = -1 Then
        lblStatus.Caption = "Select an item from history to step back to."
        Exit Sub
    End If
    
    ' Fetch target array item from the 1-based global stack index
    Dim item As Variant
    item = g_HistoryStack.item(selectedIndex + 1)
    
    Dim targetRange As Range
    Set targetRange = item(1) ' item(1) holds the Range Object
    
    ' Navigate back
    On Error Resume Next
    Application.Goto targetRange, Scroll:=True
    If Err.Number <> 0 Then
        lblStatus.Caption = "Could not navigate; worksheet may be deleted or locked."
        On Error GoTo 0
        Exit Sub
    End If
    On Error GoTo 0
    
    ' Unwind the global stack up to and including the selected index
    Dim i As Long
    For i = 0 To selectedIndex
        g_HistoryStack.Remove 1 ' Continually strip the top item
    Next i
    
    ' Sync the visual display
    RefreshListBox
    lblStatus.Caption = "Stepped back to: " & GetFriendlyAddress(targetRange)
End Sub

' --- DOUBLE CLICK LISTBOX (PEEK WITHOUT ALTERING STACK) ---
Private Sub lstHistory_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    Dim selectedIndex As Long
    selectedIndex = lstHistory.listIndex
    If selectedIndex = -1 Then Exit Sub
    
    Dim item As Variant
    item = g_HistoryStack.item(selectedIndex + 1)
    
    Dim targetRange As Range
    Set targetRange = item(1)
    
    On Error Resume Next
    Application.Goto targetRange, Scroll:=True
    If Err.Number <> 0 Then
        lblStatus.Caption = "Error peeking at range."
    Else
        lblStatus.Caption = "Peeked at: " & item(0) & " (Stack unchanged)"
    End If
    On Error GoTo 0
End Sub

' --- CLEAR ALL BUTTON ---
Private Sub btnClearAll_Click()
    Set g_HistoryStack = New Collection
    RefreshListBox
    lblStatus.Caption = "Global session history cleared."
End Sub
' --- CLEAR ALL BUTTON ---
Private Sub btnSimplifyFormula_Click()
    SimplifyIFFormulas
    ConvertNestedLookupsToDirectReferences
End Sub

' --- PRIVATE HELPER METHODS ---

Private Sub PushToGlobalStack(targetRange As Range)
    Dim displayAddress As String
    displayAddress = GetFriendlyAddress(targetRange)
    
    ' Hard structural ceiling: drop oldest entry if history exceeds 10 items
    If g_HistoryStack.Count >= 10 Then
        g_HistoryStack.Remove g_HistoryStack.Count
    End If
    
    ' Save as a 2-item array: Index 0 = address text, Index 1 = Range Object literal
    Dim stackEntry As Variant
    stackEntry = Array(displayAddress, targetRange)
    
    ' Insert directly at the top of our collection timeline (Index 1)
    If g_HistoryStack.Count = 0 Then
        g_HistoryStack.Add stackEntry
    Else
        g_HistoryStack.Add stackEntry, Before:=1
    End If
End Sub

Private Sub RefreshListBox()
    lstHistory.Clear
    If g_HistoryStack Is Nothing Then Exit Sub
    
    Dim item As Variant
    For Each item In g_HistoryStack
        lstHistory.AddItem item(0) ' Expose only the readable address text to the user
    Next item
End Sub

Private Function GetFriendlyAddress(targetRange As Range) As String
    GetFriendlyAddress = targetRange.Worksheet.Name & "!" & targetRange.Address(False, False)
End Function

Function evalStr(funcStr As String) As String
    Dim keyword As String, openParen As Long
    
    openParen = InStr(funcStr, "(")
    keyword = UCase(Trim(Left(funcStr, openParen - 1)))
    
    evalStr = funcStr
    If keyword = "VLOOKUP" Then
        Dim innerArgs As String, args() As String
        innerArgs = Mid(funcStr, openParen + 1, Len(funcStr) - openParen - 1)
        args = SplitArgs(innerArgs)
        
        If UBound(args) >= 2 Then
            Dim matchType As String: matchType = "1"
            
            If UBound(args) >= 3 Then
                Dim arg4Eval As Variant
                ' Temporarily bypass errors during 4th argument evaluation
                On Error Resume Next
                arg4Eval = Application.Evaluate(args(3))
                On Error GoTo 0
                
                If Not IsError(arg4Eval) And Not IsEmpty(arg4Eval) Then
                    If UCase(CStr(arg4Eval)) = "FALSE" Or CStr(arg4Eval) = "0" Then
                        matchType = "0"
                    End If
                End If
            End If
            
            evalStr = "INDEX(" & args(1) & ", MATCH(" & args(0) & ", INDEX(" & args(1) & ", , 1), " & matchType & "), " & args(2) & ")"
        Else
            Exit Function
        End If
    End If
End Function

Private Function SplitArgs(innerStr As String) As String()
    Dim result() As String
    ReDim result(0 To 10)
    Dim argCount As Long: argCount = 0
    Dim currentArg As String: currentArg = ""
    Dim parenCount As Long: parenCount = 0
    Dim bracketCount As Long: bracketCount = 0
    Dim inQuotes As Boolean: inQuotes = False
    Dim i As Long
    
    For i = 1 To Len(innerStr)
        Dim char As String
        char = Mid(innerStr, i, 1)
        
        If char = """" Then
            inQuotes = Not inQuotes
            currentArg = currentArg & char
        ElseIf inQuotes Then
            currentArg = currentArg & char
        Else
            If char = "(" Then
                parenCount = parenCount + 1
            ElseIf char = ")" Then
                parenCount = parenCount - 1
            ElseIf char = "[" Then
                bracketCount = bracketCount + 1
            ElseIf char = "]" Then
                bracketCount = bracketCount - 1
            End If
            
            If char = "," And parenCount = 0 And bracketCount = 0 Then
                If argCount > UBound(result) Then ReDim Preserve result(0 To argCount + 5)
                result(argCount) = Trim(currentArg)
                argCount = argCount + 1
                currentArg = ""
            Else
                currentArg = currentArg & char
            End If
        End If
    Next i
    
    If argCount > UBound(result) Then ReDim Preserve result(0 To argCount)
    result(argCount) = Trim(currentArg)
    ReDim Preserve result(0 To argCount)
    

    
    SplitArgs = result
End Function

