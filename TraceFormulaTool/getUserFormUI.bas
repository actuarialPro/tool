Attribute VB_Name = "getUserFormUI"
Public Sub ExportFormGeometryToCSV()
    Dim ctl As Control
    Dim csvLine As String
    
    ' 1. Print Header with Position and Size Column Names
    Debug.Print "Control Name,Type,Value/Caption,Left,Top,Width,Height"
    
    ' 2. Loop through every control on the UserForm
    ' (Change "UserForm1" to match your actual form's name)
    For Each ctl In frmGotoHistory.Controls
        
        On Error Resume Next
        
        ' 3. Build CSV string for Name, Type, and Value
        csvLine = Chr(34) & ctl.Name & Chr(34) & "," & _
                  Chr(34) & TypeName(ctl) & Chr(34) & ","
                  
        If Err.Number = 0 Then
            csvLine = csvLine & Chr(34) & ctl.Caption & Chr(34) & ","
        Else
            csvLine = csvLine & Chr(34) & "" & Chr(34) & ","
        End If
        
        On Error GoTo 0
        
        ' 4. Append Left, Top, Width, and Height coordinates
        csvLine = csvLine & ctl.Left & "," & _
                            ctl.Top & "," & _
                            ctl.Width & "," & _
                            ctl.Height
        
        ' 5. Output to Immediate Window
        Debug.Print csvLine
    Next ctl
End Sub
