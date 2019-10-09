Namespace XML

    ''' <summary>
    ''' XML-элемент столбца листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Class SheetColumnXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент столбца листа Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="sheetColumn">Столбец</param>
        Sub New(sheetColumn As SheetColumn)
            MyBase.New(NS.wb + "col",
                       New XAttribute("min", sheetColumn.Min),
                       New XAttribute("max", sheetColumn.Max),
                       New XAttribute("width", sheetColumn.Width),
                       New XAttribute("customWidth", 1))
        End Sub

    End Class

End Namespace

