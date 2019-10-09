Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент основного стиля в Excel для формата OfficeOpenXML.
    ''' </summary>
    Class CellStyleXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент основного стиля в Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="styleCell">Основной стиль</param>
        Sub New(styleCell As IStyleCell)
            MyBase.New(NS.wb + "cellStyle",
                       New XAttribute("name", styleCell.Name),
                       New XAttribute("xfId", styleCell.Id))

            If styleCell.Type <> 0 Then
                Throw New Exception("Неверный тип стиля ячейки.")
            End If
        End Sub

    End Class

End Namespace

