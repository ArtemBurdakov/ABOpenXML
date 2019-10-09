Namespace ContentType

    ''' <summary>
    ''' Тип контента - Книга Excel в формате OfficeOpenXML.
    ''' </summary>
    Class WorkBookContentType
        Implements IOverrideContentType

        Property PartName As String = "/xl/workbook.xml" Implements IOverrideContentType.PartName
        Property ContentType As String = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet.main+xml" Implements IOverrideContentType.ContentType
        Property Unique As Boolean = True Implements IOverrideContentType.Unique

    End Class

End Namespace

