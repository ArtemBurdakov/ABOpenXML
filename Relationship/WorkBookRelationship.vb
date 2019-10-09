Namespace Relationship

    ''' <summary>
    ''' Тип связи - Книга Excel в формате OfficeOpenXML.
    ''' </summary>
    Class WorkBookRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/officeDocument" Implements IRelationship.Type
        Property Target As String = "xl/workbook.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = True Implements IRelationship.Unique

    End Class

End Namespace

