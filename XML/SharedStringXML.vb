Namespace XML

    ''' <summary>
    ''' XML-элемент общей строки книги Excel для формата OfficeOpenXML.
    ''' </summary>
    Class SharedStringXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент общей строки книги Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="sharedString">Общая строка</param>
        Sub New(sharedString As SharedString)
            MyBase.New(NS.wb + "si",
                       New XElement(NS.wb + "t", sharedString.Value))
        End Sub

    End Class

End Namespace

