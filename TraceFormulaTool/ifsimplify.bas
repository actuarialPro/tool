Attribute VB_Name = "ifsimplify"
Option Explicit

' Set to True to watch the Outside-In onion peeling process in the Immediate Window (Ctrl + G)
Public Const DEBUG_IF_MODE As Boolean = False

Sub SimplifyIFFormulas()
    Dim cell As Range
    Dim formulaStr As String
    Dim pos As Long, searchPos As Long
    Dim fullFunc As String
    Dim replacement As String
    
    On Error GoTo CleanUp
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    
    If DEBUG_IF_MODE Then
        Debug.Print "===================================================="
        Debug.Print "STARTING IF SIMPLIFIER - Processing " & Selection.Count & " cell(s)"
        Debug.Print "===================================================="
    End If
    
    For Each cell In Selection
        If cell.HasFormula Then
            formulaStr = cell.formula
            searchPos = 1
            
            If DEBUG_IF_MODE Then
                Debug.Print ""
                Debug.Print "[CELL] Target: " & cell.Address(False, False)
                Debug.Print "[CELL] Original: " & formulaStr
            End If
            
            Do
                ' Scan Left-to-Right to catch the outermost wrapper IF statements first
                pos = FindKeywordPosForward(formulaStr, "IF", searchPos)
                
                If pos = 0 Then Exit Do
                
                If DEBUG_IF_MODE Then Debug.Print "   -> Found outer 'IF' at position: " & pos
                
                fullFunc = ExtractFullFunction(formulaStr, pos)
                If DEBUG_IF_MODE Then Debug.Print "   -> Full Block: " & fullFunc
                
                replacement = ""
                If fullFunc <> "" Then
                    replacement = ProcessIfEvaluation(fullFunc, cell)
                End If
                
                If replacement <> "" Then
                    formulaStr = Left(formulaStr, pos - 1) & replacement & Mid(formulaStr, pos + Len(fullFunc))
                    ' Reset search position to the start of the replacement block to check for further logic
                    searchPos = pos
                    If DEBUG_IF_MODE Then Debug.Print "   [PRUNE] Success! Updated Formula: " & formulaStr
                Else
                    ' If evaluation failed or errored out, skip past this keyword to avoid infinite loops
                    searchPos = pos + 2
                    If DEBUG_IF_MODE Then Debug.Print "   [SKIP] Could not safely resolve condition. Leaving block intact."
                End If
            Loop
            
            If cell.formula <> formulaStr Then
                cell.Formula2 = formulaStr
                If DEBUG_IF_MODE Then Debug.Print "[CELL] Final Output Written: " & cell.formula
            End If
        End If
    Next cell
    
    If DEBUG_IF_MODE Then Debug.Print vbCrLf & "=================== MACRO FINISHED ==================="

CleanUp:
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
End Sub

' Scans forward to find a standalone "IF(" function (ignores things like "GIFT")
Private Function FindKeywordPosForward(formula As String, keyword As String, startFrom As Long) As Long
    Dim pos As Long
    pos = startFrom
    Do
        pos = InStr(pos, formula, keyword, vbTextCompare)
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
                    FindKeywordPosForward = pos
                    Exit Function
                ElseIf c <> " " Then
                    Exit For
                End If
            Next i
        End If
        pos = pos + 1
    Loop
    FindKeywordPosForward = 0
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

' Core logical evaluator for IF arguments
Private Function ProcessIfEvaluation(funcStr As String, ctxCell As Range) As String
    Dim openParen As Long, innerStr As String, args() As String
    openParen = InStr(funcStr, "(")
    innerStr = Mid(funcStr, openParen + 1, Len(funcStr) - openParen - 1)
    
    args = SplitArgs(innerStr)
    
    ' An IF statement must have at least a condition and a True branch
    If UBound(args) < 1 Then Exit Function
    
    Dim condStr As String: condStr = args(0)
    Dim trueBranch As String: trueBranch = args(1)
    Dim falseBranch As String: falseBranch = ""
    If UBound(args) >= 2 Then falseBranch = args(2)
    
    ' Evaluate the conditional test statement natively via Excel
    Dim condResult As Variant
    On Error Resume Next
    condResult = ctxCell.Worksheet.Evaluate(condStr)
    On Error GoTo 0
    
    If Not IsError(condResult) Then
        If TypeName(condResult) = "Boolean" Then
            If DEBUG_IF_MODE Then Debug.Print "      [Eval] Condition (" & condStr & ") evaluated to: " & UCase(CStr(condResult))
            
            If condResult = True Then
                ProcessIfEvaluation = trueBranch
            Else
                If falseBranch <> "" Then
                    ProcessIfEvaluation = falseBranch
                Else
                    ProcessIfEvaluation = "FALSE" ' Excel's default behavior when False branch is omitted
                End If
            End If
        End If
    End If
End Function

Private Function SplitArgs(innerStr As String) As String()
    Dim result() As String: ReDim result(0 To 10)
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
