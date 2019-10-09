Imports OpenXML.ContentType

Namespace XML

    ''' <summary>
    ''' XML-элемент компонента пакета в формате OfficeOpenXML.
    ''' </summary>
    Class OverrideContentTypesXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент компонента пакета в файле [Content_Types].xml для формата OpenXML.
        ''' </summary>
        ''' <param name="content">Компонент пакета</param>
        Sub New(content As IOverrideContentType)
            MyBase.New(NS.ct + "Override",
                       New XAttribute("PartName", content.PartName),
                       New XAttribute("ContentType", content.ContentType))
        End Sub

    End Class

End Namespace

