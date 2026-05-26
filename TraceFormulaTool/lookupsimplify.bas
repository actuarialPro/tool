Attribute VB_Name = "lookupsimplify"
Option Explicit

' GLOBAL CONFIGURATION: Set to True to see step-by-step parsing logs in the Immediate Window (Ctrl + G)
Public Const DEBUG_MODE As Boolean = False

Sub ConvertNestedLookupsToDirectReferences()
    Dim cell As Range
    Dim formulaStr As String
    Dim pos As Long, maxStartPos As Long
    Dim foundKeyword As String
    Dim fullFunc As String
    Dim directAddress As String
    
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    If DEBUG_MODE Then
        Debug.Print "===================================================="
        Debug.Print "STARTING MACRO - Processing " & Selection.Count & " cell(s)"
        Debug.Print "===================================================="
    End If
    
    For Each cell In Selection
        If cell.HasFormula Then
            formulaStr = cell.formula
            maxStartPos = Len(formulaStr)
            
            If DEBUG_MODE Then
                Debug.Print ""
                Debug.Print "[CELL] Target: " & cell.Address(False, False)
                Debug.Print "[CELL] Original Formula: " & formulaStr
            End If
            
            Do
                Dim posV As Long, posX As Long, posI As Long, posO As Long
                
                posV = FindKeywordPos(formulaStr, "VLOOKUP", maxStartPos)
                posX = FindKeywordPos(formulaStr, "XLOOKUP", maxStartPos)
                posI = FindKeywordPos(formulaStr, "INDEX", maxStartPos)
                posO = FindKeywordPos(formulaStr, "OFFSET", maxStartPos)
                
                pos = 0
                foundKeyword = ""
                If posV > 0 Then pos = posV: foundKeyword = "VLOOKUP"
                If posX > pos Then pos = posX: foundKeyword = "XLOOKUP"
                If posI > pos Then pos = posI: foundKeyword = "INDEX"
                If posO > pos Then pos = posO: foundKeyword = "OFFSET"
                
                If pos = 0 Then
                    If DEBUG_MODE Then Debug.Print "   -> No matching lookup keywords found in current string."
                    Exit Do
                End If
                
                If DEBUG_MODE Then Debug.Print "   -> Found '" & foundKeyword & "' starting at string position: " & pos
                
                fullFunc = ExtractFullFunction(formulaStr, pos)
                If DEBUG_MODE Then Debug.Print "   -> Extracted Fragment: " & fullFunc
                
                directAddress = ""
                If fullFunc <> "" Then
                    directAddress = GetDirectAddress(fullFunc, cell)
                End If
                
                If directAddress <> "" Then
                    Dim oldFormula As String: oldFormula = formulaStr
                    formulaStr = Left(formulaStr, pos - 1) & directAddress & Mid(formulaStr, pos + Len(fullFunc))
                    maxStartPos = Len(formulaStr)
                    
                    If DEBUG_MODE Then
                        Debug.Print "   [REPLACE] Swapped fragment out successfully!"
                        Debug.Print "   [REPLACE] Updated String to: " & formulaStr
                    End If
                Else
                    If DEBUG_MODE Then Debug.Print "   [SKIP] Could not resolve a coordinate for this fragment. Skipping past it."
                    maxStartPos = pos - 1
                End If
            Loop
            
            If cell.formula <> formulaStr Then
                cell.Formula2 = formulaStr
                If DEBUG_MODE Then Debug.Print "[CELL] Written to sheet: " & cell.formula
            End If
        End If
    Next cell

    If DEBUG_MODE Then Debug.Print vbCrLf & "=================== MACRO FINISHED ==================="

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

Private Function FindKeywordPos(formula As String, keyword As String, startFrom As Long) As Long
    Dim pos As Long
    pos = startFrom
    Do
        pos = InStrRev(formula, keyword, pos, vbTextCompare)
        If pos = 0 Then Exit Do
        
        Dim validPrecedent As Boolean: validPrecedent = True
        If pos > 1 Then
            Dim preChar As String
            preChar = UCase(Mid(formula, pos - 1, 1))
            If preChar Like "[A-Z0-9_.]" Then validPrecedent = False
        End If
        
        If validPrecedent Then
            Dim i As Long
            For i = pos + Len(keyword) To Len(formula)
                Dim c As String
                c = Mid(formula, i, 1)
                If c = "(" Then
                    FindKeywordPos = pos
                    Exit Function
                ElseIf c <> " " Then
                    Exit For
                End If
                Next i
        End If
        pos = pos - 1
    Loop
    FindKeywordPos = 0
End Function

Private Function ExtractFullFunction(formula As String, startPos As Long) As String
    Dim i As Long, openParen As Long, parenCount As Long
    
    openParen = InStr(startPos, formula, "(")
    If openParen = 0 Then Exit Function
    
    parenCount = 1
    For i = openParen + 1 To Len(formula)
        Dim char As String
        char = Mid(formula, i, 1)
        If char = "(" Then
            parenCount = parenCount + 1
        ElseIf char = ")" Then
            parenCount = parenCount - 1
            If parenCount = 0 Then
                ExtractFullFunction = Mid(formula, startPos, i - startPos + 1)
                Exit Function
            End If
        End If
    Next i
End Function

Private Function GetDirectAddress(funcStr As String, ctxCell As Range) As String
    Dim evalStr As String, keyword As String, openParen As Long
    
    openParen = InStr(funcStr, "(")
    keyword = UCase(Trim(Left(funcStr, openParen - 1)))
    
    If keyword = "XLOOKUP" Or keyword = "INDEX" Or keyword = "OFFSET" Then
        evalStr = funcStr
    ElseIf keyword = "VLOOKUP" Then
        Dim innerArgs As String, args() As String
        innerArgs = Mid(funcStr, openParen + 1, Len(funcStr) - openParen - 1)
        args = SplitArgs(innerArgs)
        
        If UBound(args) >= 2 Then
            Dim matchType As String: matchType = "1"
            
            If UBound(args) >= 3 Then
                Dim arg4Eval As Variant
                ' Temporarily bypass errors during 4th argument evaluation
                On Error Resume Next
                arg4Eval = ctxCell.Worksheet.Evaluate(args(3))
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
    
    If DEBUG_MODE Then Debug.Print "      [GetDirectAddress] Formulating Background Eval: " & evalStr
    
    Dim evalResult As Variant
    On Error Resume Next
    Set evalResult = ctxCell.Worksheet.Evaluate(evalStr)
    On Error GoTo 0
    
    If Not evalResult Is Nothing Then
        If TypeOf evalResult Is Range Then
            Dim r As Range, sheetName As String
            Set r = evalResult
            sheetName = r.Worksheet.Name
            GetDirectAddress = "'" & sheetName & "'!" & r.Address(False, False)
            If DEBUG_MODE Then Debug.Print "      [GetDirectAddress] Successfully resolved object address to: " & GetDirectAddress
        Else
            If DEBUG_MODE Then Debug.Print "      [GetDirectAddress] Warning: Evaluation returned data, but it was not a valid Range mapping."
        End If
    Else
        If DEBUG_MODE Then Debug.Print "      [GetDirectAddress] Error: Excel's evaluation engine returned Nothing."
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
    
    ' Debug layout to verify argument splitting and identify syntax corruption
    If DEBUG_MODE Then
        Debug.Print "      [SplitArgs] Isolated " & (argCount + 1) & " clean function arguments:"
        Dim dIdx As Long
        For dIdx = 0 To argCount
            Debug.Print "         -> Arg(" & dIdx & "): " & result(dIdx)
        Next dIdx
    End If
    
    SplitArgs = result
End Function
