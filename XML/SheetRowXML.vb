Namespace XML

    ''' <summary>
    ''' XML-элемент строки листа Excel для формате OfficeOpenXML.
    ''' </summary>
    Class SheetRowXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент строки листа Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="sheetRow">Строка</param>
        Sub New(sheetRow As SheetRow)
            MyBase.New(NS.wb + "row",
                       New XAttribute("r", sheetRow.Number))

            If Not IsNothing(sheetRow.Height) Then
                Add(New XAttribute("ht", sheetRow.Height))
                Add(New XAttribute("customHeight", 1))
            End If

            If Not IsNothing(sheetRow.OutlineLevel) Then
                Add(New XAttribute("outlineLevel", sheetRow.OutlineLevel))
                Add(New XAttribute("hidden", "1"))
                Add(New XAttribute("collapsed", "1"))
            End If
        End Sub

    End Class

End Namespace

