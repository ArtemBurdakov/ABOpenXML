Namespace ContentType

    ''' <summary>
    ''' Тип контента - Основные свойства документа в формате OfficeOpenXML.
    ''' </summary>
    Class CorePropertiesContentType
        Implements IOverrideContentType

        Property PartName As String = "/docProps/core.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-package.core-properties+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = True Implements IOverrideContentType.Unique

    End Class

End Namespace

