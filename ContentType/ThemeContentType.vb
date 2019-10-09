Namespace ContentType

    ''' <summary>
    ''' Тип контента - Тема Excel в формате OfficeOpenXML.
    ''' </summary>
    Class ThemeContentType
        Implements IOverrideContentType

        Property PartName As String = "/xl/theme/theme.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.theme+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = False Implements IOverrideContentType.Unique

    End Class

End Namespace

