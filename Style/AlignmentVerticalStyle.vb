Namespace Style

    ''' <summary>
    ''' Варианты вертикального выранивания ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class AlignmentVerticalStyle

        ''' <summary>
        ''' Выровнивание по низу.
        ''' </summary>
        Public Shared Property Bottom As String = "bottom"

        ''' <summary>
        ''' Центрированно по высоте ячейки.
        ''' </summary>
        Public Shared Property Center As String = "center"

        ''' <summary>
        ''' Когда направлени текста горизонтальное - каждая строка текста внутри ячейки равномерно распределены по высоте ячейки.
        ''' Когда направлени текста вертикальное - каждая перенесённая строка растягивается по высоте.
        ''' </summary>
        Public Shared Property Distributed As String = "distributed"

        ''' <summary>
        ''' Когда направлени текста горизонтальное - каждая строка текста внутри ячейки равномерно распределены по высоте ячейки.
        ''' Когда направлени текста вертикальное - каждая перенесённая строка выранивается по верху.
        ''' </summary>
        Public Shared Property Justify As String = "justify"

        ''' <summary>
        ''' Выровнивание по верху.
        ''' </summary>
        Public Shared Property Top As String = "top"

    End Class

End Namespace

