''' <summary>
''' Общая строка для формата OfficeOpenXML.
''' </summary>
Class SharedString

    ''' <summary>
    ''' Создаёт объект общей строки для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="number">Номер общей строки</param>
    ''' <param name="value">Содержимое общей строки</param>
    Sub New(number As Integer, value As String)
        Me.Number = number
        Me.Value = value
    End Sub

    ''' <summary>
    ''' Номер общей строки.
    ''' </summary>
    Public Property Number As Integer

    ''' <summary>
    ''' Содержимое общей строки.
    ''' </summary>
    Public Property Value As String

End Class

