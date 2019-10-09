Namespace Relationship

    ''' <summary>
    ''' Тип связи - Настройки принтера в формате OfficeOpenXML.
    ''' </summary>
    Class PrinterSettingsRelationship
        Implements IRelationship

        Property Id As String Implements IRelationship.Id
        Property Type As String = "http://schemas.openxmlformats.org/officeDocument/2006/relationships/printerSettings" Implements IRelationship.Type
        Property Target As String = "../printerSettings/printerSettings.bin" Implements IRelationship.Target
        Property TargetMode As String = Nothing Implements IRelationship.TargetMode
        Property Unique As Boolean = False Implements IRelationship.Unique

    End Class

End Namespace

