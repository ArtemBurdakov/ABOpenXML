Imports OpenXML.Style

''' <summary>
''' Строка листа Excel для формата OfficeOpenXML.
''' </summary>
Public Class SheetRow

    Private _cells As List(Of SheetCell)

    ''' <summary>
    ''' Создаёт объект строки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New()
        _cells = New List(Of SheetCell)
    End Sub

    ''' <summary>
    ''' Создаёт объект строки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="rowNumber">Номер строки</param>
    Sub New(rowNumber As Integer)
        Me.New
        Number = rowNumber
    End Sub

    ''' <summary>
    ''' Создаёт объект строки из другой строки без копирования ячеек.
    ''' </summary>
    ''' <param name="row">Строка листа</param>
    Sub New(row As SheetRow)
        Me.New
        Height = row.Height
        Number = row.Number
        OutlineLevel = row.OutlineLevel
    End Sub

    ''' <summary>
    ''' Создаёт объект строки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="rowNumber">Номер строки</param>
    ''' <param name="cellNumber">Номер ячейки</param>
    ''' <param name="value">Содержимое ячейки</param>
    ''' <param name="height">Высота строки</param>
    Sub New(rowNumber As Integer, cellNumber As String, value As Object, Optional height As Double? = Nothing, Optional style As IStyleCell = Nothing)
        Me.New(rowNumber)
        _cells.Add(New SheetCell(cellNumber, value))
        If Not IsNothing(height) Then Me.Height = height
        If Not IsNothing(style) Then _cells.Last.Style = style
    End Sub

    ''' <summary>
    ''' Номер строки.
    ''' </summary>
    Public Property Number As Integer

    ''' <summary>
    ''' Высота строки.
    ''' </summary>
    Public Property Height As Double?

    ''' <summary>
    ''' Уровень строки в дереве.
    ''' </summary>
    Public Property OutlineLevel As Short?

    ''' <summary>
    ''' Добавляет ячейку.
    ''' </summary>
    ''' <param name="cell"></param>
    Public Sub AddCell(cell As SheetCell)
        _cells.Add(cell)
    End Sub

    ''' <summary>
    ''' Возвращает ячейку по указанному индексу из строки.
    ''' </summary>
    ''' <param name="index">Индекс</param>
    Public Function CellAt(index As Integer) As SheetCell
        Return _cells.ElementAt(index)
    End Function

    ''' <summary>
    ''' Возвращает список ячеек строки.
    ''' </summary>
    Public Function GetCells() As List(Of SheetCell)
        Return _cells
    End Function

End Class

