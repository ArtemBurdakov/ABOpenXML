Namespace Style

    ''' <summary>
    ''' Форматирование цифр ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class NumFmtStyle
        Implements IEquatable(Of NumFmtStyle)

        Public Shared Operator =(obj1 As NumFmtStyle, obj2 As NumFmtStyle) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As NumFmtStyle, obj2 As NumFmtStyle) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Создаёт объект Форматирование цифр ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()

        End Sub

        ''' <summary>
        ''' Создаёт объект Форматирование цифр ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="formatCode">
        ''' Код форматирования.
        ''' 0 - обязательная цифра;
        ''' # - не обязательная цифра;
        ''' % - приведение к процентам;
        ''' , - заменяет 1000;
        ''' * - заполняет ячейку до упора символами справа звёздочки;
        ''' "text" - текст;
        ''' m - 1-12 месяцы;
        ''' mm - 01-12 месяцы;
        ''' d - 1-31 дни;
        ''' dd - 01-31 дни;
        ''' yy - 00-99 годы;
        ''' yyyy - 0000-9999 годы;
        ''' пример с временем: hh:mm:ss.00 - 04:36:03.75.
        ''' </param>
        Sub New(formatCode As String)
            Me.FormatCode = formatCode
        End Sub

        ''' <summary>
        ''' Идентификатор форматирования цифр.
        ''' </summary>
        Public Property Id As Integer

        ''' <summary>
        ''' Код форматирования.
        ''' 0 - обязательная цифра;
        ''' # - не обязательная цифра;
        ''' % - приведение к процентам;
        ''' , - заменяет 1000;
        ''' * - заполняет ячейку до упора символами справа звёздочки;
        ''' "text" - текст;
        ''' m - 1-12 месяцы;
        ''' mm - 01-12 месяцы;
        ''' d - 1-31 дни;
        ''' dd - 01-31 дни;
        ''' yy - 00-99 годы;
        ''' yyyy - 0000-9999 годы;
        ''' пример с временем: hh:mm:ss.00 - 04:36:03.75.
        ''' </summary>
        Public Property FormatCode As String

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As NumFmtStyle) As Boolean Implements IEquatable(Of NumFmtStyle).Equals
            Return other IsNot Nothing AndAlso FormatCode = other.FormatCode
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, NumFmtStyle))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return FormatCode.GetHashCode()
        End Function

    End Class

End Namespace

