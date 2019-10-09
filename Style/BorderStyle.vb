Namespace Style

    ''' <summary>
    ''' Граница ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class BorderStyle
        Implements IEquatable(Of BorderStyle)

        Public Shared Operator =(obj1 As BorderStyle, obj2 As BorderStyle) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As BorderStyle, obj2 As BorderStyle) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Создаёт объект границы ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()

        End Sub

        ''' <summary>
        ''' Идентификатор границы.
        ''' </summary>
        Public Property Id As Integer

        ''' <summary>
        ''' Стиль левой границы.
        ''' </summary>
        Public Property LeftStyle As String

        ''' <summary>
        ''' Цвет левой границы.
        ''' </summary>
        Public Property LeftColor As Color

        ''' <summary>
        ''' Стиль правой границы.
        ''' </summary>
        Public Property RightStyle As String

        ''' <summary>
        ''' Цвет правой границы.
        ''' </summary>
        Public Property RightColor As Color

        ''' <summary>
        ''' Стиль верхней границы.
        ''' </summary>
        Public Property TopStyle As String

        ''' <summary>
        ''' Цвет верхней границы.
        ''' </summary>
        Public Property TopColor As Color

        ''' <summary>
        ''' Стиль нижней границы.
        ''' </summary>
        Public Property BottomStyle As String

        ''' <summary>
        ''' Цвет нижней границы.
        ''' </summary>
        Public Property BottomColor As Color

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As BorderStyle) As Boolean Implements IEquatable(Of BorderStyle).Equals
            Return other IsNot Nothing AndAlso LeftStyle = other.LeftStyle AndAlso LeftColor = other.LeftColor AndAlso
                RightStyle = other.RightStyle AndAlso RightColor = other.RightColor AndAlso
                TopStyle = other.TopStyle AndAlso TopColor = other.TopColor AndAlso
                BottomStyle = other.BottomStyle AndAlso BottomColor = other.BottomColor
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, BorderStyle))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (LeftStyle & LeftColor.ToString & RightStyle & RightColor.ToString &
                TopStyle & TopColor.ToString & BottomStyle & BottomColor.ToString).GetHashCode()
        End Function

    End Class

End Namespace

