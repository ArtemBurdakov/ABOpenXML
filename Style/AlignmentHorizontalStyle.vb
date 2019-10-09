Namespace Style

    ''' <summary>
    ''' Варианты горизонтального выранивания ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class AlignmentHorizontalStyle

        ''' <summary>
        ''' Текст центрируется по ячейке.
        ''' </summary>
        Public Shared Property Center As String = "center"

        ''' <summary>
        ''' Текст центрируется по нескольким ячейкам.
        ''' </summary>
        Public Shared Property CenterContinuous As String = "centerContinuous"

        ''' <summary>
        ''' Каждая перенесённая строка растягивается по ширине ячейки.
        ''' Если есть отступ, то он делаестся слева и справа.
        ''' </summary>
        Public Shared Property Distributed As String = "distributed"

        ''' <summary>
        ''' Текст должен быть заполнен по всей ширине ячейки.
        ''' </summary>
        Public Shared Property Fill As String = "fill"

        ''' <summary>
        ''' Текст выравнивается по левому краю, чилса и даты по правому, булевы типы по центру.
        ''' </summary>
        Public Shared Property General As String = "general"

        ''' <summary>
        ''' Каждая перенесённая строка выравнивается по левому краю.
        ''' </summary>
        Public Shared Property Justify As String = "justify"

        ''' <summary>
        ''' Содержимое выравнивается по левому краю.
        ''' </summary>
        Public Shared Property Left As String = "left"

        ''' <summary>
        ''' Содержимое выравнивается по правому краю.
        ''' </summary>
        Public Shared Property Right As String = "right"

    End Class

End Namespace

