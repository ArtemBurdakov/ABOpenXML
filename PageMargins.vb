Imports OpenXML.XML

Public Class PageMargins

    Public Property Footer As Double
    Public Property Header As Double
    Public Property Bottom As Double
    Public Property Top As Double
    Public Property Right As Double
    Public Property Left As Double

    Public Function GetXElement() As XElement
        Return New XElement(NS.wb + "pageMargins",
                                 New XAttribute("footer", Footer),
                                 New XAttribute("header", Header),
                                 New XAttribute("bottom", Bottom),
                                 New XAttribute("top", Top),
                                 New XAttribute("right", Right),
                                 New XAttribute("left", Left))
    End Function

End Class

