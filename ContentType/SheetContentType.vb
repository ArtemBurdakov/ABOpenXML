Namespace ContentType

    ''' <summary>
    ''' Тип контента - Лист Excel в формате OfficeOpenXML.
    ''' </summary>
    Class SheetContentType
        Implements IOverrideContentType

        Property PartName As String = "/xl/worksheets/sheet.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.spreadsheetml.worksheet+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = False Implements IOverrideContentType.Unique

    End Class

End Namespace

