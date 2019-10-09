Namespace ContentType

    ''' <summary>
    ''' Тип контента - XML-компонент пакета в формате OfficeOpenXML.
    ''' </summary>
    Class XmlContentType
        Implements IDefaultContentType

        Property Extension As String = "xml" Implements IDefaultContentType.Extension
        Property ContentType As String = "application/xml" Implements IDefaultContentType.ContentType

    End Class

End Namespace

