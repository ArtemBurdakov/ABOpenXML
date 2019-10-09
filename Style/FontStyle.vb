Namespace Style

    ''' <summary>
    ''' Шрифт текста ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class FontStyle
        Implements IEquatable(Of FontStyle)

        Public Shared Operator =(obj1 As FontStyle, obj2 As FontStyle) As Boolean
            Return (obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing)
        End Operator

        Public Shared Operator <>(obj1 As FontStyle, obj2 As FontStyle) As Boolean
            Return Not ((obj1 IsNot Nothing AndAlso obj1.Equals(obj2)) OrElse (obj1 Is Nothing And obj2 Is Nothing))
        End Operator

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        Sub New()
            Bold = False
            Italic = False
            Underline = False
            Size = 11
            Color = New Color With {.RGB = "FF000000"}
            Name = "Calibri"
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="font">Другой шрифт</param>
        Sub New(font As FontStyle)
            Bold = font.Bold
            Color = font.Color
            Italic = font.Italic
            Name = font.Name
            Size = font.Size
            Underline = font.Underline
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        Sub New(name As String)
            Me.New
            Me.Name = name
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="size">Размер шрифта</param>
        Sub New(size As Integer)
            Me.New
            Me.Size = size
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="size">Размер шрифта</param>
        Sub New(name As String, size As Integer)
            Me.New(name)
            Me.Size = size
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="color">Цвет шрифта в формате ARGB (значение от 00 до FF)</param>
        Sub New(name As String, color As String)
            Me.New(name)
            Me.Color = New Color With {.RGB = color}
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="size">Размер шрифта</param>
        ''' <param name="color">Цвет шрифта в формате ARGB (значение от 00 до FF)</param>
        Sub New(size As Integer, color As String)
            Me.New
            Me.Size = size
            Me.Color = New Color With {.RGB = color}
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="size">Размер шрифта</param>
        ''' <param name="color">Цвет шрифта в формате ARGB (значение от 00 до FF)</param>
        Sub New(name As String, size As Integer, color As String)
            Me.New(name)
            Me.Size = size
            Me.Color = New Color With {.RGB = color}
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="bold">Отображение текста полужырным</param>
        ''' <param name="italic">Отображение текста курсивом</param>
        ''' <param name="underline">Отображение текста подчёркнутым</param>
        Sub New(name As String, bold As Boolean, italic As Boolean, underline As Boolean)
            Me.New(name)
            Me.Bold = bold
            Me.Italic = italic
            Me.Underline = underline
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="size">Размер шрифта</param>
        ''' <param name="bold">Отображение текста полужырным</param>
        ''' <param name="italic">Отображение текста курсивом</param>
        ''' <param name="underline">Отображение текста подчёркнутым</param>
        Sub New(size As Integer, bold As Boolean, italic As Boolean, underline As Boolean)
            Me.New(size)
            Me.Bold = bold
            Me.Italic = italic
            Me.Underline = underline
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="color">Цвет шрифта в формате ARGB (значение от 00 до FF)</param>
        ''' <param name="bold">Отображение текста полужырным</param>
        ''' <param name="italic">Отображение текста курсивом</param>
        ''' <param name="underline">Отображение текста подчёркнутым</param>
        Sub New(name As String, color As String, bold As Boolean, italic As Boolean, underline As Boolean)
            Me.New(name, color)
            Me.Bold = bold
            Me.Italic = italic
            Me.Underline = underline
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="size">Размер шрифта</param>
        ''' <param name="bold">Отображение текста полужырным</param>
        ''' <param name="italic">Отображение текста курсивом</param>
        ''' <param name="underline">Отображение текста подчёркнутым</param>
        Sub New(name As String, size As Integer, bold As Boolean, italic As Boolean, underline As Boolean)
            Me.New(name, size)
            Me.Bold = bold
            Me.Italic = italic
            Me.Underline = underline
        End Sub

        ''' <summary>
        ''' Создаёт объект шрифта текста ячейки в Excel для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Название шрифта</param>
        ''' <param name="size">Размер шрифта</param>
        ''' <param name="color">Цвет шрифта в формате ARGB (значение от 00 до FF)</param>
        ''' <param name="bold">Отображение текста полужырным</param>
        ''' <param name="italic">Отображение текста курсивом</param>
        ''' <param name="underline">Отображение текста подчёркнутым</param>
        Sub New(name As String, size As Integer, color As String, bold As Boolean, italic As Boolean, underline As Boolean)
            Me.New(name, size, color)
            Me.Bold = bold
            Me.Italic = italic
            Me.Underline = underline
        End Sub

        ''' <summary>
        ''' Идентификатор шрифта.
        ''' </summary>
        Public Property Id As Integer

        ''' <summary>
        ''' Отображает текст полужырным.
        ''' </summary>
        Public Property Bold As Boolean

        ''' <summary>
        ''' Отображает текст курсивом.
        ''' </summary>
        Public Property Italic As Boolean

        ''' <summary>
        ''' Отображает текст подчёркнутым.
        ''' </summary>
        Public Property Underline As Boolean

        ''' <summary>
        ''' Размер шрифта.
        ''' </summary>
        Public Property Size As Integer

        ''' <summary>
        ''' Цвет шрифта в формате ARGB (значение от 00 до FF).
        ''' </summary>
        Public Property Color As Color

        ''' <summary>
        ''' Название шрифта.
        ''' </summary>
        Public Property Name As String

        ''' <summary>
        ''' Определяет равен ли объект, текущему объекту.
        ''' </summary>
        Public Overloads Function Equals(other As FontStyle) As Boolean Implements IEquatable(Of FontStyle).Equals
            Return other IsNot Nothing AndAlso Bold = other.Bold AndAlso Italic = other.Italic AndAlso
                Underline = other.Underline AndAlso Size = other.Size AndAlso Color = other.Color AndAlso
                Name = other.Name
        End Function

        Public Overrides Function Equals(obj As Object) As Boolean
            Return Equals(DirectCast(obj, FontStyle))
        End Function

        Public Overrides Function GetHashCode() As Integer
            Return (Bold.ToString & Italic.ToString & Underline.ToString &
                Size.ToString & Color.ToString & Name).GetHashCode()
        End Function

    End Class

End Namespace

