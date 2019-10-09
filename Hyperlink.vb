Imports OpenXML.XML

''' <summary>
''' Представляет гиперссылку.
''' </summary>
Public Class Hyperlink

    ''' <summary>
    ''' Идентификатор связи, относящейся к этой гиперссылки.
    ''' </summary>
    Friend Property Id As String

    ''' <summary>
    ''' Номер ячейки в которой находится ссылка.
    ''' </summary>
    Public Property NumberCell As String

    ''' <summary>
    ''' Возвращает xml-элемент.
    ''' </summary>
    Friend Function GetXElement() As XElement
        Return New XElement(NS.wb + "hyperlink",
            New XAttribute("ref", NumberCell),
            New XAttribute(NS.r + "id", Id))
    End Function

End Class

