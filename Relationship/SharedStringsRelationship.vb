Namespace Relationship

    ''' <summary>
    ''' Тип связи - Общие строки в Excel в формате OfficeOpenXML.
    ''' </summary>
    Class SharedStringsRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/sharedStrings" Implements IRelationship.Type
        Property Target As String = "sharedStrings.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = True Implements IRelationship.Unique

    End Class

End Namespace

