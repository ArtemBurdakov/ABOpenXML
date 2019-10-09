Imports OpenXML.XML

Public Class PageSetup

    Public Property Orientation As String
    Public Property Scale As Integer?

    Function GetXElement() As XElement
        Dim pageSetup = New XElement(NS.wb + "pageSetup")

        If Orientation IsNot Nothing Then
            pageSetup.Add(New XAttribute("orientation", Orientation))
        End If

        pageSetup.Add(New XAttribute("paperSize", 9))

        If Scale IsNot Nothing Then
            pageSetup.Add(New XAttribute("scale", Scale))
        End If

        Return pageSetup
    End Function

End Class

