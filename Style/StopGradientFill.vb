Namespace Style

    ''' <summary>
    ''' Точка остоновки градиента заливки ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StopGradientFill
        Implements IEquatable(Of StopGradientFill)

        Public Shared Operator =(obj1 As StopGradientFill, obj2 As StopGradientFill) As Boolean
            Return obj1.Equals(obj2)
        End Operator

        Public Shared Operator <>(obj1 As StopGradientFill, obj2 As StopGradientFill) As Boolean
            Return Not obj1.Equals(obj2)
        End Operator

        ''' <summary>
        ''' Создаёт объект точку остоновки градиента заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="color">Цвет в формате ARGB (значение от 00 до FF)</param>
        ''' <param name="position">Место начала чистого цвета в процентном отношении. Этот атрибут ограничен значениями от 0 до 1.</param>
        Sub New(color As String, position As Double)
            If position > 1 Then Throw New Exception("Значение позиции точки остоновки градиента выше 1.")
            Me.Position = position
            Me.Color = New Color With {.RGB = color}
        End Sub

        ''' <summary>
        ''' Место начала чистого цвета в процентном отношении.
        ''' Этот атрибут ограничен значениями от 0 до 1.
        ''' </summary>
        Public Property Position As Double

        ''' <summary>
        ''' Цвет в формате ARGB (значение от 00 до FF).
        ''' </summary>
        Public Property Color As Color

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As StopGradientFill) As Boolean Implements IEquatable(Of StopGradientFill).Equals
            Dim i = -1
            Return other IsNot Nothing AndAlso Position = other.Position AndAlso Color = other.Color
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, StopGradientFill))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return ToString.GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Return Position & Color.ToString
        End Function

    End Class

End Namespace

