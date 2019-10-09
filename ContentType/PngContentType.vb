Namespace ContentType

    ''' <summary>
    ''' Тип контента - Изображенние формата png в формате OfficeOpenXML.
    ''' </summary>
    Class PngContentType
        Implements IDefaultContentType

        Property Extension As String = "png" Implements IDefaultContentType.Extension
        Property ContentType As String = "image/png" Implements IDefaultContentType.ContentType

    End Class

End Namespace

