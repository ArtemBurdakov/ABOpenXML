Namespace Relationship

    ''' <summary>
    ''' Тип связи - Дополнительные свойства документа в формате OfficeOpenXML.
    ''' </summary>
    Class ExtendedPropertiesRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/extended-properties" Implements IRelationship.Type
        Property Target As String = "docProps/app.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = True Implements IRelationship.Unique

    End Class

End Namespace

