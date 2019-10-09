Namespace Style

    ''' <summary>
    ''' Градиент заливки ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class GradientFill
        Implements IEquatable(Of GradientFill)

        Private _stops As List(Of StopGradientFill)

        Public Shared Operator =(obj1 As GradientFill, obj2 As GradientFill) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As GradientFill, obj2 As GradientFill) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Создаёт объект градиента заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()
            _stops = New List(Of StopGradientFill)
        End Sub

        ''' <summary>
        ''' Создаёт объект градиента заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="degree">Угол линейного градиента</param>
        Sub New(degree As Integer)
            Me.New()
            Type = 0
            Me.Degree = degree
        End Sub

        ''' <summary>
        ''' Создаёт объект градиента заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="bottom">Указывает в процентном формате положение нижнего края внутреннего прямоугольника. Этот атрибут ограничен значениями от 0 до 1</param>
        ''' <param name="left">Указывает в процентном формате положение левого края внутреннего прямоугольника. Этот атрибут ограничен значениями от 0 до 1</param>
        ''' <param name="right">Указывает в процентном формате положение правого края внутреннего прямоугольника. Этот атрибут ограничен значениями от 0 до 1</param>
        ''' <param name="top">Указывает в процентном формате положение верхнего края внутреннего прямоугольника. Этот атрибут ограничен значениями от 0 до 1</param>
        Sub New(bottom As Double, left As Double, right As Double, top As Double)
            Me.New()
            Type = 1
            Me.Bottom = bottom
            Me.Left = left
            Me.Right = right
            Me.Top = top
        End Sub

        ''' <summary>
        ''' Тип градиента: 0 - линейный, 1 - прямоугольный.
        ''' </summary>
        Public Property Type As Integer

        ''' <summary>
        ''' Угол линейного градиента.
        ''' </summary>
        Public Property Degree As Integer

        ''' <summary>
        ''' Указывает в процентном формате положение нижнего края внутреннего прямоугольника.
        ''' Этот атрибут ограничен значениями от 0 до 1.
        ''' </summary>
        Public Property Bottom As Double

        ''' <summary>
        ''' Указывает в процентном формате положение левого края внутреннего прямоугольника.
        ''' Этот атрибут ограничен значениями от 0 до 1.
        ''' </summary>
        Public Property Left As Double

        ''' <summary>
        ''' Указывает в процентном формате положение правого края внутреннего прямоугольника.
        ''' Этот атрибут ограничен значениями от 0 до 1.
        ''' </summary>
        Public Property Right As Double

        ''' <summary>
        ''' Указывает в процентном формате положение верхнего края внутреннего прямоугольника.
        ''' Этот атрибут ограничен значениями от 0 до 1.
        ''' </summary>
        Public Property Top As Double

        ''' <summary>
        ''' Возвращает список точек изменения цвета.
        ''' </summary>
        Public Function GetStops() As List(Of StopGradientFill)
            Return _stops
        End Function

        ''' <summary>
        ''' Добавляет точку изменения цвета
        ''' </summary>
        ''' <param name="color">Цвет в формате ARGB (значение от 00 до FF)</param>
        ''' <param name="position">Место начала чистого цвета в процентном отношении. Этот атрибут ограничен значениями от 0 до 1.</param>
        Public Sub AddStop(color As String, position As Double)
            _stops.Add(New StopGradientFill(color, position))
        End Sub

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As GradientFill) As Boolean Implements IEquatable(Of GradientFill).Equals
            Dim i = -1
            Return other IsNot Nothing AndAlso Type = other.Type AndAlso Degree = other.Degree AndAlso
                Bottom = other.Bottom AndAlso Left = other.Left AndAlso Right = other.Right AndAlso Top = other.Top AndAlso
                _stops.TrueForAll(Function(x)
                                      i += 1
                                      If x = other.GetStops.ElementAt(i) Then
                                          Return True
                                      End If
                                      Return False
                                  End Function)
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, GradientFill))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return ToString.GetHashCode()
        End Function

        Public Overrides Function ToString() As String
            Dim stops = ""
            For Each s In _stops
                stops = s.ToString
            Next
            Return Type & Degree & Bottom & Left & Right & Top & stops
        End Function

    End Class

End Namespace

