Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент шрифта в Excel для формата OfficeOpenXML.
    ''' </summary>
    Class FontStyleXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент шрифта в Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="fontStyle">Шрифт</param>
        Sub New(fontStyle As FontStyle)
            MyBase.New(NS.wb + "font")

            If fontStyle.Bold Then
                Add(New XElement(NS.wb + "b"))
            End If

            If fontStyle.Italic Then
                Add(New XElement(NS.wb + "i"))
            End If

            If fontStyle.Underline Then
                Add(New XElement(NS.wb + "u"))
            End If

            Add(New XElement(NS.wb + "sz",
                New XAttribute("val", fontStyle.Size)))
            Add(fontStyle.Color.GetXElement)
            Add(New XElement(NS.wb + "name",
                New XAttribute("val", fontStyle.Name)))

        End Sub

    End Class

End Namespace

