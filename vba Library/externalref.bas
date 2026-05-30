Attribute VB_Name = "externalref"
' ==============================================================================
' FUNCTION: ExtractPathsFromFormula
' PURPOSE: Uses Regular Expressions to extract external workbook/sheet paths
'          from a cell formula, returning an array of unique external references.
' ==============================================================================
Function ExtractPathsFromFormula(formula As String) As Variant
    Dim regex As Object
    Dim matches As Object
    Dim match As Object
    Dim paths As Object
    Dim pathArray() As String
    Dim i As Long
    
    ' Create RegExp object
    Set regex = CreateObject("VBScript.RegExp")
    regex.Pattern = "'?(([A-Z]:\\[^']*)?\[([^\]]+)\]([^'!]+))'?"
    regex.Global = True
    
    ' Use Dictionary to avoid duplicates
    Set paths = CreateObject("Scripting.Dictionary")
    
    ' Find all matches
    Set matches = regex.Execute(formula)
    
    For Each match In matches
        If match.SubMatches.count > 0 Then
            If Not paths.Exists(match.SubMatches(0)) Then
                paths.Add match.SubMatches(0), Nothing
            End If
        End If
    Next match
    
    ' Convert Dictionary keys to array
    If paths.count > 0 Then
        ReDim pathArray(0 To paths.count - 1)
        For i = 0 To paths.count - 1
            pathArray(i) = paths.Keys()(i)
        Next i
        ExtractPathsFromFormula = pathArray
    Else
        ExtractPathsFromFormula = Array() ' Return empty array
    End If
End Function

' ==============================================================================
' SUB: ShowAndOpenExternalReferences
' SHORTCUT: Ctrl + Shift + Q
' PURPOSE: Checks the active cell for external links, displays them in an
'          InputBox, and opens/navigates to the user's selection.
' ==============================================================================
Sub ShowAndOpenExternalReferences()
    Dim userInput As String
    Dim selectedIndex As Integer
    Dim i As Long
    Dim selectedPath As String
    Dim filepath As String
    Dim fileName As String
    Dim SheetName As String
    Dim wb As Workbook
    Dim bracketStart As Long
    Dim bracketEnd As Long
    Dim formula As String
    Dim paths As Variant
    Dim displayText As String
    
    ' Check if active cell has a formula
    If Not ActiveCell.HasFormula Then
        MsgBox "The active cell does not contain a formula.", vbExclamation
        Exit Sub
    End If
    
    ' Get the formula
    formula = ActiveCell.formula
    
    ' Extract paths
    paths = ExtractPathsFromFormula(formula)
    
    ' Check if any paths were found
    If Not IsArray(paths) Or UBound(paths) < 0 Then
        MsgBox "No external references found in this formula.", vbInformation
        Exit Sub
    End If
    
    ' Build display text
    displayText = "External references found:" & vbCrLf & vbCrLf
    For i = LBound(paths) To UBound(paths)
        displayText = displayText & (i + 1) & ". " & paths(i) & vbCrLf & vbCrLf
    Next i
    displayText = displayText & vbCrLf & "Enter the number to open the file and sheet:"
    
    ' Show input box
    userInput = InputBox(displayText, "Select External Reference", 1)
    
    ' Check if user cancelled
    If userInput = "" Then Exit Sub
    
    ' Validate input
    If Not IsNumeric(userInput) Then
        MsgBox "Please enter a valid number.", vbExclamation
        Exit Sub
    End If
    
    selectedIndex = CInt(userInput) - 1
    
    ' Check if index is valid
    If selectedIndex < LBound(paths) Or selectedIndex > UBound(paths) Then
        MsgBox "Invalid selection. Please enter a number between 1 and " & (UBound(paths) + 1), vbExclamation
        Exit Sub
    End If
    
    ' Get selected path
    selectedPath = paths(selectedIndex)
    
    ' Parse the path to extract file path and sheet name
    ' Format: C:\path\[workbook.ext]SheetName
    bracketStart = InStr(selectedPath, "[")
    bracketEnd = InStr(selectedPath, "]")
    
    If bracketStart = 0 Or bracketEnd = 0 Then
        MsgBox "Unable to parse the file path.", vbExclamation
        Exit Sub
    End If
    
    ' Extract components safely
    filepath = Left(selectedPath, bracketStart - 1) & Mid(selectedPath, bracketStart + 1, bracketEnd - bracketStart - 1)
    fileName = Mid(selectedPath, bracketStart + 1, bracketEnd - bracketStart - 1)
    SheetName = Mid(selectedPath, bracketEnd + 1)
    
    ' Open the workbook
    If InStr(selectedPath, "\") > 0 Then
        On Error Resume Next
        Set wb = Workbooks.Open(filepath, addtomru:=True)
        On Error GoTo 0
        
        If wb Is Nothing Then
            MsgBox "Unable to open the file: " & vbCrLf & filepath, vbExclamation
            Exit Sub
        End If
    Else
        ' If it's already open or has no directory path info
        Set wb = Workbooks(fileName)
        wb.Activate
        Application.WindowState = xlMaximized
    End If
    
    ' Activate the sheet
    On Error Resume Next
    wb.Worksheets(SheetName).Activate
    If Err.Number <> 0 Then
        Application.StatusBar = "Sheet '" & SheetName & "' not found in the workbook."
        Err.Clear
    Else
        Application.StatusBar = "opened: " & filepath
    End If
    On Error GoTo 0
    
End Sub
