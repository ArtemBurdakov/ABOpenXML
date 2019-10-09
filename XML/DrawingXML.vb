Namespace XML

    ''' <summary>
    ''' XML-элемент графики для формата OfficeOpenXML.
    ''' </summary>
    Class DrawingXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент графики для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="drawing">Графика</param>
        Sub New(drawing As Drawing)
            MyBase.New(NS.wb + "drawing",
                       New XAttribute(NS.r + "id", drawing.Id))
        End Sub

    End Class

End Namespace

