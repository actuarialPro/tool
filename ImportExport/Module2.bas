Attribute VB_Name = "Module2"
Sub ExtractLinksShort()
    Dim links As Variant
    links = ThisWorkbook.LinkSources(xlExcelLinks)
    
    Range("B2:B99").ClearContents
    For i = 1 To UBound(links)
        Range("B" & i + 1).Value = links(i)
        i = i + 1
    Next i
End Sub
Sub UpdateLinks()
    Dim i As Long

    For i = 1 To 99
        If .Range("A" & i).Value = "Yes" And .Range("C" & i).Value <> "" Then
            ThisWorkbook.ChangeLink .Range("B" & i).Value, .Range("C" & i).Value, xlExcelLinks
        End If
    Next i

End Sub
