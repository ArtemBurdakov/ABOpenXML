Namespace Style

    ''' <summary>
    ''' Стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleCell
        Implements IStyleCell

        ''' <summary>
        ''' Создаёт объект стиля ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()
            Type = 1
            ParentId = 0
            Id = 0
            Name = "Default"
            Font = New FontStyle
            Fill = New FillStyle
            Border = New BorderStyle
        End Sub

        ''' <summary>
        ''' Создаёт объект стиля ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="style">Другой стиль</param>
        Sub New(style As IStyleCell)
            Type = 1
            ParentId = 0
            Id = 0
            Name = style.Name
            Alignment = New AlignmentStyle(style.Alignment)
            Border = style.Border
            Font = New FontStyle(style.Font)
            Fill = style.Fill
            NumFmt = style.NumFmt
        End Sub

        ''' <summary>
        ''' Создаёт объект стиля ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="type">Тип стиля</param>
        ''' <param name="id">Идентификатор стиля</param>
        ''' <param name="parentId">Идентификатор родителя</param>
        ''' <param name="name">Наименование стиля</param>
        Sub New(type As Integer, id As Integer, parentId As Integer, name As String)
            Me.Id = id
            Me.Type = type
            Me.ParentId = parentId
            Me.Name = name
            Font = New FontStyle
            Fill = New FillStyle
            Border = New BorderStyle
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

End Namespace

