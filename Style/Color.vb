Imports OpenXML.XML

Namespace Style

    ''' <summary>
    ''' Представляет цвет для формата OfficeOpenXML.
    ''' </summary>
    Public Class Color
        Implements IEquatable(Of Color)

        Public Shared Operator =(obj1 As Color, obj2 As Color) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As Color, obj2 As Color) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Цвет в формате AARRGGBB.
        ''' </summary>
        Public RGB As String

        ''' <summary>
        ''' Номер темы (заданный цвет).
        ''' </summary>
        Public Theme As Integer

        ''' <summary>
        ''' Возвращает XML-элемент данного объекта.
        ''' </summary>
        Friend Function GetXElement() As XElement
            Dim result = New XElement(NS.wb + "color")
            If RGB IsNot Nothing Then
                result.Add(New XAttribute("rgb", RGB))
            Else
                result.Add(New XAttribute("theme", Theme))
            End If
            Return result
        End Function

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As Color) As Boolean Implements IEquatable(Of Color).Equals
            Return other IsNot Nothing AndAlso RGB = other.RGB AndAlso Theme = other.Theme
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, Color))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (RGB & Theme).GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return RGB & Theme
        End Function

    End Class

End Namespace

