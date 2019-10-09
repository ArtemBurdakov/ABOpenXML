Namespace Relationship

    ''' <summary>
    ''' Тип связи - Тема в Excel в формате OfficeOpenXML.
    ''' </summary>
    Class ThemeRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/theme" Implements IRelationship.Type
        Property Target As String = "theme/theme.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = False Implements IRelationship.Unique

    End Class

End Namespace

