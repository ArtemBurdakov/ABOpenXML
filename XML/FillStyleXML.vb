Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент заливки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Class FillStyleXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент заливки в Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="fillStyle">Заливка</param>
        Sub New(fillStyle As FillStyle)
            MyBase.New(NS.wb + "fill")

            If Not IsNothing(fillStyle.PatternType) Then
                Add(New XElement(NS.wb + "patternFill",
                                 New XAttribute("patternType", fillStyle.PatternType)))

                If Not IsNothing(fillStyle.ForegroundColor) Then
                    Dim fgColor = fillStyle.ForegroundColor.GetXElement
                    fgColor.Name = NS.wb + "fgColor"
                    Element(NS.wb + "patternFill").Add(fgColor)
                End If

                If Not IsNothing(fillStyle.BackgroundColor) Then
                    Dim bgColor = fillStyle.BackgroundColor.GetXElement
                    bgColor.Name = NS.wb + "bgColor"
                    Element(NS.wb + "patternFill").Add(bgColor)
                End If
            End If

            If Not IsNothing(fillStyle.GradientFill) Then
                Add(New XElement(NS.wb + "gradientFill"))

                If fillStyle.GradientFill.Type = 0 Then
                    Dim gf = Element(NS.wb + "gradientFill")
                    gf.Add(New XAttribute("degree", fillStyle.GradientFill.Degree))

                    For Each s In fillStyle.GradientFill.GetStops
                        gf.Add(New XElement(NS.wb + "stop",
                                            New XAttribute("position", s.Position),
                                            s.Color.GetXElement))
                    Next
                ElseIf fillStyle.GradientFill.Type = 1 Then
                    Dim gf = Element(NS.wb + "gradientFill")
                    gf.Add(New XAttribute("type", "path"),
                           New XAttribute("left", fillStyle.GradientFill.Left),
                           New XAttribute("right", fillStyle.GradientFill.Right),
                           New XAttribute("top", fillStyle.GradientFill.Top),
                           New XAttribute("bottom", fillStyle.GradientFill.Bottom))

                    For Each s In fillStyle.GradientFill.GetStops
                        gf.Add(New XElement(NS.wb + "stop",
                                            New XAttribute("position", s.Position),
                                            s.Color.GetXElement))
                    Next
                End If
            End If
        End Sub

    End Class

End Namespace

