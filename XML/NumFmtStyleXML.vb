Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент форматирования цифр в Excel для формата OfficeOpenXML.
    ''' </summary>
    Class NumFmtStyleXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент форматирования цифр в Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="numFmtStyle">Форматирование цифр</param>
        Sub New(numFmtStyle As NumFmtStyle)
            MyBase.New(NS.wb + "numFmt",
                       New XAttribute("numFmtId", numFmtStyle.Id),
                       New XAttribute("formatCode", If(IsNothing(numFmtStyle.FormatCode), "", numFmtStyle.FormatCode)))
        End Sub

    End Class

End Namespace

