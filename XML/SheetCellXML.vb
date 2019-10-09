Namespace XML

    ''' <summary>
    ''' XML-элемент ячейки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Class SheetCellXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент ячейки листа Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="sheetCell">Ячейка</param>
        Sub New(sheetCell As SheetCell)
            MyBase.New(NS.wb + "c",
                       New XAttribute("r", sheetCell.Number),
                       New XElement(NS.wb + "v", sheetCell.Value))

            If Not IsNothing(sheetCell.Type) Then
                Add(New XAttribute("t", sheetCell.Type))
            End If

            If Not IsNothing(sheetCell.Style) Then
                Add(New XAttribute("s", sheetCell.Style.Id))
            End If

        End Sub

    End Class

End Namespace

