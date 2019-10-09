Namespace XML

    ''' <summary>
    ''' XML-элемент якоря привязки контента графики для формата OfficeOpenXML.
    ''' </summary>
    Class AnchorXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент якоря привязки контента графики для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="anchor">Якорь привязки контента графики</param>
        Sub New(anchor As IAnchor)
            MyBase.New(NS.xdr + anchor.Type)

            If anchor.Type = "twoCellAnchor" Then
                Dim twoCellAnchor = CType(anchor, TwoCellAnchor)
                Add(New XElement(NS.xdr + "from",
                                    New XElement(NS.xdr + "col", twoCellAnchor.FromCol),
                                    New XElement(NS.xdr + "colOff", twoCellAnchor.FromColOff),
                                    New XElement(NS.xdr + "row", twoCellAnchor.FromRow),
                                    New XElement(NS.xdr + "rowOff", twoCellAnchor.FromRowOff)),
                       New XElement(NS.xdr + "to",
                                    New XElement(NS.xdr + "col", twoCellAnchor.ToCol),
                                    New XElement(NS.xdr + "colOff", twoCellAnchor.ToColOff),
                                    New XElement(NS.xdr + "row", twoCellAnchor.ToRow),
                                    New XElement(NS.xdr + "rowOff", twoCellAnchor.ToRowOff)),
                    New XElement(NS.xdr + "clientData"))
            End If

        End Sub

    End Class

End Namespace

