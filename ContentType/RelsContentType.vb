Namespace ContentType

    ''' <summary>
    ''' Тип контента - Описатель связей в формате OfficeOpenXML.
    ''' </summary>
    Class RelsContentType
        Implements IDefaultContentType

        Property Extension As String = "rels" Implements IDefaultContentType.Extension
        Property ContentType As String = "application/vnd.openxmlformats-package.relationships+xml" Implements IDefaultContentType.ContentType

    End Class

End Namespace

