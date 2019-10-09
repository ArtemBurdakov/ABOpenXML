Namespace Style

    ''' <summary>
    ''' Интерфейс стиля ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Interface IStyleCell

        Property Type As Integer
        Property Id As Integer
        Property ParentId As Integer
        Property Name As String
        Property Alignment As AlignmentStyle
        Property Font As FontStyle
        Property Fill As FillStyle
        Property Border As BorderStyle
        Property NumFmt As NumFmtStyle

    End Interface


End Namespace

