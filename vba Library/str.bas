Attribute VB_Name = "str"


' Simple f-string style formatter for VBA using {} placeholders
Function FStr(template As String, ParamArray args() As Variant) As String
    Dim i As Long
    Dim result As String
    result = template
    
    On Error GoTo ErrHandler
    
    ' Replace each {} with the corresponding argument
    For i = LBound(args) To UBound(args)
        result = Replace(result, "{}", CStr(args(i)), , 1) ' Replace only first occurrence each time
    Next i
    
    FStr = result
    Exit Function
    
ErrHandler:
    FStr = "#FStr Error: " & Err.Description
End Function

' Example usage
Sub TestFStr()
    Dim name As String
    Dim age As Integer
    
    name = "Alice"
    age = 30
    
    Debug.Print FStr("My name is {}, I am {} years old.", name, age)
    Debug.Print FStr("Coordinates: ({}, {})", 12.34, 56.78)
    Debug.Print RTextafter("abc\def\th", "\")
    Debug.Print lfind("abc\def\th", "h")
End Sub

