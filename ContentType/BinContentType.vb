Namespace ContentType

    ''' <summary>
    ''' Тип контента - Настройки печати в формате OfficeOpenXML.
    ''' </summary>
    Class BinContentType
        Implements IDefaultContentType

        Property Extension As String = "bin" Implements IDefaultContentType.Extension
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.spreadsheetml.printerSettings" Implements IDefaultContentType.ContentType

    End Class

End Namespace

