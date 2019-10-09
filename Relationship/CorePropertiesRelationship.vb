Namespace Relationship

    ''' <summary>
    ''' Тип связи - Основные свойства документа в формате OfficeOpenXML.
    ''' </summary>
    Class CorePropertiesRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/package/2006/relationships/metadata/core-properties" Implements IRelationship.Type
        Property Target As String = "docProps/core.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = True Implements IRelationship.Unique

    End Class

End Namespace

