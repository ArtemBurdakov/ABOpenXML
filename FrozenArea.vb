Imports OpenXML.XML

Public Class FrozenArea

    Private _topLeftCell As String

    Public Property XSplit As Integer
    Public Property YSplit As Integer

    Public Function GetXElement() As XElement
        Dim frozenArea =
            New XElement(NS.wb + "sheetViews",
                New XElement(NS.wb + "sheetView", New XAttribute("tabSelected", 1),
                             New XAttribute("workbookViewId", 0),
                    New XElement(NS.wb + "pane", New XAttribute("xSplit", XSplit),
                                 New XAttribute("ySplit", YSplit), New XAttribute("topLeftCell", GetTopLeftCell()),
                                 New XAttribute("state", "frozen"))))

        Return frozenArea
    End Function

    Private Function GetTopLeftCell() As String
        Dim numberColumn = ""

        If XSplit = 0 Then
            numberColumn = "A"
        ElseIf XSplit = 1 Then
            numberColumn = "B"
        ElseIf XSplit = 2 Then
            numberColumn = "C"
        ElseIf XSplit = 3 Then
            numberColumn = "D"
        ElseIf XSplit = 4 Then
            numberColumn = "E"
        ElseIf XSplit = 5 Then
            numberColumn = "F"
        ElseIf XSplit = 6 Then
            numberColumn = "G"
        ElseIf XSplit = 7 Then
            numberColumn = "H"
        ElseIf XSplit = 8 Then
            numberColumn = "I"
        ElseIf XSplit = 9 Then
            numberColumn = "J"
        ElseIf XSplit = 10 Then
            numberColumn = "K"
        ElseIf XSplit = 11 Then
            numberColumn = "L"
        ElseIf XSplit = 12 Then
            numberColumn = "M"
        ElseIf XSplit = 13 Then
            numberColumn = "N"
        ElseIf XSplit = 14 Then
            numberColumn = "O"
        ElseIf XSplit = 15 Then
            numberColumn = "P"
        End If

        Return numberColumn & (YSplit + 1)
    End Function

End Class

