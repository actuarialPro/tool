Attribute VB_Name = "bookmark"
Option Explicit

' Define the Struct (User-Defined Type)
Public Type BookmarkType
    WorkbookName As String
    SheetName As String
    Address As String
End Type

' Global list storing the structures natively
Public BookmarkHistory(1 To 9) As BookmarkType
Public BookmarkGetMode As Boolean
Public IsSystemInitialized As Boolean

' Ensures the structured list is clean on first load
Public Sub InitializeBookmarkSystem()
    If Not IsSystemInitialized Then
        Dim i As Long
        For i = 1 To 9
            BookmarkHistory(i).WorkbookName = ""
            BookmarkHistory(i).SheetName = ""
            BookmarkHistory(i).Address = ""
        Next i
        IsSystemInitialized = True
    End If
End Sub

' Entry point macro to launch the application panel
Sub ShowBookmarkManagerGetMode()
Attribute ShowBookmarkManagerGetMode.VB_ProcData.VB_Invoke_Func = "K\n14"
    BookmarkGetMode = True
    Bookmarks.Show
End Sub

' Entry point macro to launch the application panel
Sub ShowBookmarkManagerSetMode()
Attribute ShowBookmarkManagerSetMode.VB_ProcData.VB_Invoke_Func = "J\n14"
    BookmarkGetMode = False
    Bookmarks.Show
End Sub
