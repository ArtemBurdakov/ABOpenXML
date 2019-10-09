Namespace Style

    ''' <summary>
    ''' Заливка ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class FillStyle
        Implements IEquatable(Of FillStyle)

        Public Shared Operator =(obj1 As FillStyle, obj2 As FillStyle) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As FillStyle, obj2 As FillStyle) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Создаёт объект заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()
            PatternType = PatternTypeStyle.None
        End Sub

        ''' <summary>
        ''' Создаёт объект заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="patternType">Тип узора</param>
        Sub New(patternType As String)
            Me.PatternType = patternType
        End Sub

        ''' <summary>
        ''' Создаёт объект заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="patternType">Тип узора</param>
        ''' <param name="foregroundColor">Цвет переднего плана</param>
        Sub New(patternType As String, foregroundColor As String)
            Me.PatternType = patternType
            Me.ForegroundColor = New Color With {.RGB = foregroundColor}
        End Sub

        ''' <summary>
        ''' Создаёт объект заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="patternType">Тип узора</param>
        ''' <param name="foregroundColor">Цвет переднего плана</param>
        ''' <param name="backgroundColor">Цвет фона</param>
        Sub New(patternType As String, foregroundColor As String, backgroundColor As String)
            Me.New(patternType, foregroundColor)
            Me.BackgroundColor = New Color With {.RGB = backgroundColor}
        End Sub

        ''' <summary>
        ''' Создаёт объект заливки ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="gradientFill">Градиент заливки</param>
        Sub New(gradientFill As GradientFill)
            Me.GradientFill = gradientFill
        End Sub

        ''' <summary>
        ''' Идентификатор заливки.
        ''' </summary>
        Public Property Id As Integer

        ''' <summary>
        ''' Тип узора.
        ''' </summary>
        Public Property PatternType As String

        ''' <summary>
        ''' Цвет переднего плана.
        ''' </summary>
        Public Property ForegroundColor As Color

        ''' <summary>
        ''' Цвет фона.
        ''' </summary>
        Public Property BackgroundColor As Color

        ''' <summary>
        ''' Градиент заливки.
        ''' </summary>
        Public Property GradientFill As GradientFill

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As FillStyle) As Boolean Implements IEquatable(Of FillStyle).Equals
            Return other IsNot Nothing AndAlso PatternType = other.PatternType AndAlso ForegroundColor = other.ForegroundColor AndAlso
                 BackgroundColor = other.BackgroundColor AndAlso GradientFill = other.GradientFill
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, FillStyle))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (PatternType & ForegroundColor.ToString & BackgroundColor.ToString & GradientFill.ToString).GetHashCode()
        End Function

    End Class

End Namespace

