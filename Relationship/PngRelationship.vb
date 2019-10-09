Namespace Relationship

    ''' <summary>
    ''' Тип связи - Изображенние формата png в формате OfficeOpenXML.
    ''' </summary>
    Class PngRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String =
            "http://schemas.openxmlformats.org/officeDocument/2006/relationships/image" Implements IRelationship.Type
        Property Target As String = "../media/image.png" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = False Implements IRelationship.Unique

    End Class

End Namespace

