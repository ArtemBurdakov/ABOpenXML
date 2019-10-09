Namespace Style

    ''' <summary>
    ''' Выранивание ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class AlignmentStyle
        Implements IEquatable(Of AlignmentStyle)

        Public Shared Operator =(obj1 As AlignmentStyle, obj2 As AlignmentStyle) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As AlignmentStyle, obj2 As AlignmentStyle) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Создаёт объект выранивание ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()

        End Sub

        ''' <summary>
        ''' Создаёт объект выранивание ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="alignment">Объект выранивания</param>
        Sub New(alignment As AlignmentStyle)
            Horizontal = alignment.Horizontal
            Indent = alignment.Indent
            JustifyLastLine = alignment.JustifyLastLine
            ReadingOrder = alignment.ReadingOrder
            ShrinkToFit = alignment.ShrinkToFit
            TextRotation = alignment.TextRotation
            Vertical = alignment.Vertical
            WrapText = alignment.WrapText
        End Sub

        ''' <summary>
        ''' Создаёт объект выранивание ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="horizontal">Горизонтальное выранивание</param>
        Sub New(horizontal As String)
            Me.Horizontal = horizontal
        End Sub

        ''' <summary>
        ''' Создаёт объект выранивание ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="horizontal">Горизонтальное выранивание</param>
        ''' <param name="wrapText">Логическое значение, указывающее, должен ли текст переноситься по словам</param>
        Sub New(horizontal As String, wrapText As Boolean)
            Me.Horizontal = horizontal
            Me.WrapText = wrapText
        End Sub

        ''' <summary>
        ''' Создаёт объект выранивание ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="horizontal">Горизонтальное выранивание</param>
        ''' <param name="vertical">Вертикальное выравнивание</param>
        Sub New(horizontal As String, vertical As String)
            Me.Horizontal = horizontal
            Me.Vertical = vertical
        End Sub

        ''' <summary>
        ''' Создаёт объект выранивание ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="horizontal">Горизонтальное выранивание</param>
        ''' <param name="vertical">Вертикальное выравнивание</param>
        ''' <param name="wrapText">Логическое значение, указывающее, должен ли текст переноситься по словам</param>
        Sub New(horizontal As String, vertical As String, wrapText As Boolean)
            Me.New(horizontal, vertical)
            Me.WrapText = wrapText
        End Sub

        ''' <summary>
        ''' Идентификатор выравнивания.
        ''' </summary>
        Public Property Id As Integer

        ''' <summary>
        ''' Горизонтальное выранивание.
        ''' </summary>
        Public Property Horizontal As String

        ''' <summary>
        ''' Отступ. Кол-во пробелов равно значению умноженному на 3.
        ''' </summary>
        Public Property Indent As Integer?

        ''' <summary>
        ''' Логическое значение, указывающее, должны ли ячейки выравниваться или распределяться.
        ''' </summary>
        Public Property JustifyLastLine As Boolean?

        ''' <summary>
        ''' Направление чтение текста: 0 - зависимость контекста, 1 - слева направо, 2 - справа налево.
        ''' </summary>
        Public Property ReadingOrder As Integer?

        ''' <summary>
        ''' Логическое значение, указывающее, следует ли сокращать отображаемый текст в ячейке, чтобы соответствовать ширине ячейки.
        ''' Не применимо, если ячейка содержит несколько строк текста.
        ''' </summary>
        Public Property ShrinkToFit As Boolean?

        ''' <summary>
        ''' Вращение текста в ячейке. Выражается в градусах. Значения находятся в диапазоне от 0 до 180.
        ''' Первая буква текста считается центральной точкой дуги.
        ''' </summary>
        Public Property TextRotation As Integer?

        ''' <summary>
        ''' Вертикальное выравнивание.
        ''' </summary>
        Public Property Vertical As String

        ''' <summary>
        ''' Логическое значение, указывающее, должен ли текст переноситься по словам.
        ''' </summary>
        Public Property WrapText As Boolean?

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As AlignmentStyle) As Boolean Implements IEquatable(Of AlignmentStyle).Equals
            Return other IsNot Nothing AndAlso Horizontal = other.Horizontal AndAlso Vertical = other.Vertical AndAlso
                ((Indent IsNot Nothing AndAlso other.Indent IsNot Nothing AndAlso Indent = other.Indent) OrElse
                (Indent Is Nothing AndAlso other.Indent Is Nothing)) AndAlso
                ((JustifyLastLine IsNot Nothing AndAlso other.JustifyLastLine IsNot Nothing AndAlso
                JustifyLastLine = other.JustifyLastLine) OrElse
                (JustifyLastLine Is Nothing AndAlso other.JustifyLastLine Is Nothing)) AndAlso
                ((ReadingOrder IsNot Nothing AndAlso other.ReadingOrder IsNot Nothing AndAlso
                ReadingOrder = other.ReadingOrder) OrElse
                (ReadingOrder Is Nothing AndAlso other.ReadingOrder Is Nothing)) AndAlso
                ((ShrinkToFit IsNot Nothing AndAlso other.ShrinkToFit IsNot Nothing AndAlso
                ShrinkToFit = other.ShrinkToFit) OrElse
                (ShrinkToFit Is Nothing AndAlso other.ShrinkToFit Is Nothing)) AndAlso
                ((TextRotation IsNot Nothing AndAlso other.TextRotation IsNot Nothing AndAlso
                TextRotation = other.TextRotation) OrElse
                (TextRotation Is Nothing AndAlso other.TextRotation Is Nothing)) AndAlso
                ((WrapText IsNot Nothing AndAlso other.WrapText IsNot Nothing AndAlso WrapText = other.WrapText) OrElse
                (WrapText Is Nothing AndAlso other.WrapText Is Nothing))
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, AlignmentStyle))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (Horizontal & Indent.ToString & JustifyLastLine.ToString & TextRotation.ToString &
                ReadingOrder.ToString & ShrinkToFit.ToString & Vertical & WrapText.ToString).GetHashCode()
        End Function

    End Class

End Namespace

