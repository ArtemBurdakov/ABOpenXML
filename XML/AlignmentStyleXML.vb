Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент выравнивание в стиле Excel для формата OfficeOpenXML.
    ''' </summary>
    Class AlignmentStyleXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент выравнивания в стиле Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="alignmentStyle">Выравнивание</param>
        Sub New(alignmentStyle As AlignmentStyle)
            MyBase.New(NS.wb + "alignment")

            If Not IsNothing(alignmentStyle.Horizontal) Then
                Add(New XAttribute("horizontal", alignmentStyle.Horizontal))
            End If

            If Not IsNothing(alignmentStyle.Indent) Then
                Add(New XAttribute("indent", alignmentStyle.Indent))
            End If

            If Not IsNothing(alignmentStyle.JustifyLastLine) Then
                Add(New XAttribute("justifyLastLine", If(alignmentStyle.JustifyLastLine, 1, 0)))
            End If

            If Not IsNothing(alignmentStyle.ReadingOrder) Then
                Add(New XAttribute("readingOrder", alignmentStyle.ReadingOrder))
            End If

            If Not IsNothing(alignmentStyle.ShrinkToFit) Then
                Add(New XAttribute("shrinkToFit", If(alignmentStyle.ShrinkToFit, 1, 0)))
            End If

            If Not IsNothing(alignmentStyle.TextRotation) Then
                Add(New XAttribute("textRotation", alignmentStyle.TextRotation))
            End If

            If Not IsNothing(alignmentStyle.Vertical) Then
                Add(New XAttribute("vertical", alignmentStyle.Vertical))
            End If

            If Not IsNothing(alignmentStyle.WrapText) Then
                Add(New XAttribute("wrapText", If(alignmentStyle.WrapText, 1, 0)))
            End If

        End Sub

    End Class

End Namespace

