Namespace XML

    ''' <summary>
    ''' XML-элемент листа Excel для формате OfficeOpenXML.
    ''' </summary>
    Class SheetXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент листа Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="sheet">Лист</param>
        Sub New(sheet As Sheet)
            MyBase.New(NS.wb + "sheet",
                       New XAttribute("name", sheet.Name),
                       New XAttribute("sheetId", sheet.SheetId),
                       New XAttribute(NS.r + "id", sheet.Id))
        End Sub

    End Class

End Namespace

