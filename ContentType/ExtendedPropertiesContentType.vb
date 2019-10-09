Namespace ContentType

    ''' <summary>
    ''' Тип контента - Дополнительные свойства документа в формате OfficeOpenXML.
    ''' </summary>
    Class ExtendedPropertiesContentType
        Implements IOverrideContentType

        Property PartName As String = "/docProps/app.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.extended-properties+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = True Implements IOverrideContentType.Unique

    End Class

End Namespace

