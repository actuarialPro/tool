Attribute VB_Name = "WorkbookSelector"
Attribute VB_Base = "0{6C2BBA6F-34E7-4CB7-9A48-20154CBEAA23}{5561E111-1B25-4A3C-9A9A-0553DFC69D9E}"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Attribute VB_TemplateDerived = False
Attribute VB_Customizable = False
Option Explicit
Private MasterList As Object
Private SelectedWorkbook As Workbook

Private Sub lblTitle_Click()

End Sub

Private Sub UserForm_Initialize()
    Dim wb As Workbook
    Dim i As Long
    Set SelectedWorkbook = Nothing
    Set MasterList = CreateObject("Scripting.Dictionary")
    For i = 1 To Application.Workbooks.count
        Set MasterList(Application.Workbooks(i).name) = Application.Workbooks(i)
    Next i
    PopulateList MasterList
    txtFilter.SetFocus
End Sub

Private Sub txtFilter_Change()
    Dim temp As Object
    Dim key As Variant
    Dim filter As String
    Set temp = CreateObject("Scripting.Dictionary")
    filter = LCase(Trim(txtFilter.text))
    If Len(filter) = 0 Then
        PopulateList MasterList
        Exit Sub
    End If
    For Each key In MasterList.Keys
        If InStr(1, LCase(key), filter, vbTextCompare) > 0 Then
            Set temp(key) = MasterList(key)
        End If
    Next
    PopulateList temp
End Sub

Private Sub PopulateList(ByVal dict As Object)
    Dim key As Variant
    lstworkbooks.Clear
    Dim counter As Integer
    For Each key In dict.Keys
        lstworkbooks.AddItem key
        If ActiveWorkbook.name = key Then lstworkbooks.ListIndex = counter
        counter = counter + 1
    Next key

End Sub

Private Sub cmdOK_Click()
    If lstworkbooks.ListIndex < 0 Then
        MsgBox "Please select a workbook.", vbExclamation, "No Selection"
        Exit Sub
    End If
    Dim selectedName As String
    selectedName = lstworkbooks.List(lstworkbooks.ListIndex)
    If Not MasterList.Exists(selectedName) Then
        MsgBox "Selected workbook is no longer available.", vbExclamation
        Exit Sub
    End If
    Set SelectedWorkbook = MasterList(selectedName)
    SelectedWorkbook.Activate
    Unload Me
End Sub

Private Sub cmdCancel_Click()
    Unload Me
End Sub

Private Sub lstworkbooks_DblClick(ByVal Cancel As MSForms.ReturnBoolean)
    cmdOK_Click
End Sub

Private Sub lstworkbooks_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        cmdOK_Click
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub txtFilter_KeyDown(ByVal KeyCode As MSForms.ReturnInteger, ByVal Shift As Integer)
    If KeyCode = vbKeyReturn Then
        KeyCode = 0
        If lstworkbooks.ListIndex >= 0 Then
            cmdOK_Click
        End If
    ElseIf KeyCode = vbKeyUp Then
        KeyCode = 0
        If lstworkbooks.ListIndex > 0 Then
            lstworkbooks.ListIndex = lstworkbooks.ListIndex - 1
        End If
    ElseIf KeyCode = vbKeyDown Then
        KeyCode = 0
        If lstworkbooks.ListIndex < lstworkbooks.ListCount - 1 Then
            lstworkbooks.ListIndex = lstworkbooks.ListIndex + 1
        End If
    ElseIf KeyCode = vbKeyEscape Then
        KeyCode = 0
        Unload Me
    End If
End Sub

Private Sub UserForm_Terminate()
    Set MasterList = Nothing
    Set SelectedWorkbook = Nothing
End Sub

