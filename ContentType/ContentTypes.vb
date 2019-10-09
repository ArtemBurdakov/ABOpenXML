Imports OpenXML.Zip
Imports OpenXML.XML

Namespace ContentType

    ''' <summary>
    ''' Список типов контенета для формата OfficeOpenXML.
    ''' </summary>
    Public Class ContentTypes

        Private _doc As XDocument

        ''' <summary>
        ''' Список типов контенета для формата OpenXML.
        ''' </summary>
        Sub New()
            _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                                    {New XElement(NS.ct + "Types")})
        End Sub

        ''' <summary>
        ''' Возвращает XML файл.
        ''' </summary>
        Public Function GetXML() As XDocument
            Return _doc
        End Function

        ''' <summary>
        ''' Добавляет в Zip-архив файл [Content_Types].xml.
        ''' </summary>
        ''' <param name="zip">Zip-архив</param>
        ''' <returns>Возвращает ZIP-архив</returns>
        Public Function AddZip(zip As ZipFile) As ZipFile
            zip.CreateFile("[Content_Types].xml", _doc)
            Return zip
        End Function

        ''' <summary>
        ''' Добавляет тип содержимого или компонент пакета.
        ''' </summary>
        ''' <param name="content">Тип содержимого</param>
        Public Sub Add(content As IDefaultContentType)
            If _doc.Root.Elements(NS.ct + "Default") _
                   .Where(Function(x) x.Attribute("Extension").Value = content.Extension).Count = 0 Then
                _doc.Root.Add(New DefaultContentTypesXML(content))
            End If
        End Sub

        ''' <summary>
        ''' Добавляет тип содержимого или компонент пакета.
        ''' </summary>
        ''' <param name="content">Компонент пакета</param>
        Public Sub Add(content As IOverrideContentType)
            Dim contents = _doc.Root.Elements(NS.ct + "Override") _
                   .Where(Function(x) x.Attribute("ContentType").Value = content.ContentType)

            If contents.Count = 0 And content.Unique Then
                _doc.Root.Add(New OverrideContentTypesXML(content))
            ElseIf Not content.Unique Then
                Dim indexDot = content.PartName.LastIndexOf("."c)
                content.PartName = content.PartName.Substring(0, indexDot) & (contents.Count + 1) & content.PartName.Substring(indexDot)
                _doc.Root.Add(New OverrideContentTypesXML(content))
            End If
        End Sub

    End Class

End Namespace

