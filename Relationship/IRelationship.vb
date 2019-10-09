Namespace Relationship

    ''' <summary>
    ''' Интерфейс связи в формате OfficeOpenXML.
    ''' </summary>
    Interface IRelationship

        Property Id As String
        Property Type As String
        Property Target As String
        Property TargetMode As String
        Property Unique As Boolean

    End Interface

End Namespace

