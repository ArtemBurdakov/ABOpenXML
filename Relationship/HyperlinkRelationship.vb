Namespace Relationship

    ''' <summary>
    ''' Тип связи - Гиперссылка в формате OfficeOpenXML.
    ''' </summary>
    Class HyperlinkRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/hyperlink" Implements IRelationship.Type
        Property Target As String = "" Implements IRelationship.Target
        Property TargetMode As String = "External" Implements IRelationship.TargetMode
        Property Unique As Boolean = False Implements IRelationship.Unique

    End Class

End Namespace

