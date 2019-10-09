Namespace XML

    ''' <summary>
    ''' XML-элемент контента графики для формата OfficeOpenXML.
    ''' </summary>
    Class ContentDrawingXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент контента графики для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="content">Контент графики</param>
        Sub New(content As IContentDrawing)
            MyBase.New(New AnchorXML(content.Anchor))
            Element(NS.xdr + "clientData").AddBeforeSelf(New XElement(NS.xdr + content.Type))

            If content.Type = "pic" Then
                Element(NS.xdr + content.Type).Add(New XElement(NS.xdr + "nvPicPr",
                                                                New XElement(NS.xdr + "cNvPr",
                                                                             New XAttribute("id", content.IdContent),
                                                                             New XAttribute("name", content.Name)),
                                                                New XElement(NS.xdr + "cNvPicPr",
                                                                             New XElement(NS.a + "picLocks",
                                                                                          New XAttribute("noChangeAspect", 1)))),
                                                   New XElement(NS.xdr + "blipFill",
                                                                New XElement(NS.a + "blip",
                                                                             New XAttribute(NS.r + "embed", content.Id)),
                                                                New XElement(NS.a + "stretch",
                                                                             New XElement(NS.a + "fillRect"))),
                                                   New XElement(NS.xdr + "spPr",
                                                                New XElement(NS.a + "prstGeom",
                                                                             New XAttribute("prst", "rect"))))
            End If
        End Sub

    End Class

End Namespace

