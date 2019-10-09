Imports OpenXML.Style

''' <summary>
''' Ячейка листа Excel для формата OfficeOpenXML.
''' </summary>
Public Class SheetCell

    ''' <summary>
    ''' Создаёт объект ячейки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New(number As String, value As Object)
        Me.Number = number
        Me.Value = value
        If TypeOf value Is String Then Type = "s"
    End Sub

    ''' <summary>
    ''' Создаёт объект ячейки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New(number As String, value As Object, style As IStyleCell)
        Me.New(number, value)
        Me.Style = style
    End Sub

    ''' <summary>
    ''' Создаёт объект ячейки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New(number As String, value As Date, style As IStyleCell)
        Me.New(number, (value.Year - 1900) * 365 + Math.Floor((value.Year - 1901) / 4) -
                Math.Floor((value.Year - 1901) / 100) + Math.Floor((value.Year - 1601) / 400) +
                value.DayOfYear + 1)
        Me.Style = style
    End Sub

    ''' <summary>
    ''' Номер ячейки.
    ''' </summary>
    Public Property Number As String

    ''' <summary>
    ''' Содержимое ячейки.
    ''' </summary>
    Public Property Value As Object

    ''' <summary>
    ''' Тип ячейки.
    ''' </summary>
    Public Property Type As String

    ''' <summary>
    ''' Стиль ячейки.
    ''' </summary>
    Public Property Style As IStyleCell

End Class

