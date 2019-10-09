Namespace ContentType

    ''' <summary>
    ''' Тип контента - Стили Excel в формате OfficeOpenXML.
    ''' </summary>
    Class StylesContentType
        Implements IOverrideContentType

        Property PartName As String = "/xl/styles.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.spreadsheetml.styles+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = True Implements IOverrideContentType.Unique

    End Class

End Namespace

