Namespace Relationship

    ''' <summary>
    ''' Тип связи - Лист Excel в формате OfficeOpenXML.
    ''' </summary>
    Class SheetRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/worksheet" Implements IRelationship.Type
        Property Target As String = "worksheets/sheet.xml" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = False Implements IRelationship.Unique

    End Class

End Namespace

