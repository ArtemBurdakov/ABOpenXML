''' <summary>
''' Стобец листа Excel для формата OfficeOpenXML.
''' </summary>
Class SheetColumn

    ''' <summary>
    ''' Создаёт объект стобца листа Excel для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="min"></param>
    ''' <param name="max"></param>
    ''' <param name="width"></param>
    Sub New(min As Integer, max As Integer, width As Double)
        Me.Min = min
        Me.Max = max
        Me.Width = width
    End Sub

    ''' <summary>
    ''' Начальный порядковый номер столбца
    ''' </summary>
    Public Property Min As Integer

    ''' <summary>
    ''' Конечный порядковый номер столбца
    ''' </summary>
    Public Property Max As Integer

    ''' <summary>
    ''' Ширина столбца
    ''' </summary>
    Public Property Width As Double

End Class
