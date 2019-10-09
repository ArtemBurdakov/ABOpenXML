Namespace XML

    ''' <summary>
    ''' XML-элемент объединённой ячейки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Class MergeCellXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент объединённой ячейки листа Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="mergeCell">Объеденённая ячейка</param>
        Sub New(mergeCell As MergeCell)
            MyBase.New(NS.wb + "mergeCell",
                       New XAttribute("ref", mergeCell.Ref))
        End Sub

    End Class

End Namespace

