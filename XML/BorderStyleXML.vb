Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент границы в Excel для формата OfficeOpenXML.
    ''' </summary>
    Class BorderStyleXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент границы в Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="borderStyle">Граница</param>
        Sub New(borderStyle As BorderStyle)
            MyBase.New(NS.wb + "border")

            If Not IsNothing(borderStyle.LeftStyle) Then
                Add(New XElement(NS.wb + "left",
                                 New XAttribute("style", borderStyle.LeftStyle)))

                If Not IsNothing(borderStyle.LeftColor) Then
                    Element(NS.wb + "left").Add(borderStyle.LeftColor.GetXElement)
                End If
            End If

            If Not IsNothing(borderStyle.RightStyle) Then
                Add(New XElement(NS.wb + "right",
                                 New XAttribute("style", borderStyle.RightStyle)))

                If Not IsNothing(borderStyle.RightColor) Then
                    Element(NS.wb + "right").Add(borderStyle.RightColor.GetXElement)
                End If
            End If

            If Not IsNothing(borderStyle.TopStyle) Then
                Add(New XElement(NS.wb + "top",
                                 New XAttribute("style", borderStyle.TopStyle)))

                If Not IsNothing(borderStyle.TopColor) Then
                    Element(NS.wb + "top").Add(borderStyle.TopColor.GetXElement)
                End If
            End If

            If Not IsNothing(borderStyle.BottomStyle) Then
                Add(New XElement(NS.wb + "bottom",
                                 New XAttribute("style", borderStyle.BottomStyle)))

                If Not IsNothing(borderStyle.BottomColor) Then
                    Element(NS.wb + "bottom").Add(borderStyle.BottomColor.GetXElement)
                End If
            End If

        End Sub

    End Class

End Namespace

