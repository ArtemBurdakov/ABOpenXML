Imports OpenXML.ContentType

Namespace XML

    ''' <summary>
    ''' XML-элемент стандартного типа содержимого в формате OfficeOpenXML.
    ''' </summary>
    Class DefaultContentTypesXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент стандартного типа содержимого в файле [Content_Types].xml для формата OpenXML.
        ''' </summary>
        ''' <param name="сontent">Стандартный тип содержимого</param>
        Sub New(сontent As IDefaultContentType)
            MyBase.New(NS.ct + "Default",
                       New XAttribute("Extension", сontent.Extension),
                       New XAttribute("ContentType", сontent.ContentType))
        End Sub

    End Class

End Namespace

