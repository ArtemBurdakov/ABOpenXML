Namespace Relationship

    ''' <summary>
    ''' Тип связи - Стили в Excel в формате OfficeOpenXML.
    ''' </summary>
    Class StylesRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/styles" Implements IRelationship.Type
        Property Target As String = "styles.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = True Implements IRelationship.Unique

    End Class

End Namespace

