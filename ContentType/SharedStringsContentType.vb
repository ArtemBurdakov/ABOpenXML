Namespace ContentType

    ''' <summary>
    ''' Тип контента - Общие строки в Excel в формате OfficeOpenXML.
    ''' </summary>
    Class SharedStringsContentType
        Implements IOverrideContentType

        Property PartName As String = "/xl/sharedStrings.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.spreadsheetml.sharedStrings+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = True Implements IOverrideContentType.Unique

    End Class

End Namespace

