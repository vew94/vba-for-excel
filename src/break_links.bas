Option Explicit

'/**
' * Break links
' */
Sub break_links()
    call fnc_break_links(ThisWorkbook)

End Sub

Private Function fnc_break_links(target as Workbook)
    Dim source as Variant
    For Each source in target.LinkSources(Type:=xlLinkTypeExcelLinks)
        target.BreakLink Name:=source, Type:=xlLinkTypeExcelLinks
    next source

End Function