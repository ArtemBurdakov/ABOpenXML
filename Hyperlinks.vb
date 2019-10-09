Imports OpenXML.XML

''' <summary>
''' Представляет список гиперссылок.
''' </summary>
Public Class Hyperlinks
    Implements IEnumerable(Of Hyperlink)

    Private _links As New List(Of Hyperlink)

    ''' <summary>
    ''' Добавляет гиперссылку в конец списка.
    ''' </summary>
    ''' <param name="link">Гиперссылка</param>
    Public Sub Add(link As Hyperlink)
        _links.Add(link)
    End Sub

    ''' <summary>
    ''' Возвращает xml-элемент.
    ''' </summary>
    Friend Function GetXElement() As XElement
        Dim result = New XElement(NS.wb + "hyperlinks")
        For Each l In _links
            result.Add(l.GetXElement)
        Next
        Return result
    End Function

    Public Function GetEnumerator() As IEnumerator(Of Hyperlink) Implements IEnumerable(Of Hyperlink).GetEnumerator
        Return DirectCast(_links, IEnumerable(Of Hyperlink)).GetEnumerator()
    End Function

    Private Function IEnumerable_GetEnumerator() As IEnumerator Implements IEnumerable.GetEnumerator
        Return DirectCast(_links, IEnumerable(Of Hyperlink)).GetEnumerator()
    End Function

End Class

