Namespace ContentType

    ''' <summary>
    ''' Тип контента - Рисунок в Excel в формате OfficeOpenXML.
    ''' </summary>
    Class DrawingContentType
        Implements IOverrideContentType

        Property PartName As String = "/xl/drawings/drawing.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.drawing+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = False Implements IOverrideContentType.Unique

    End Class

End Namespace

