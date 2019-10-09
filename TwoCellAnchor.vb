''' <summary>
''' Якорь привязки контента графики для формата OfficeOpenXML.
''' </summary>
Public Class TwoCellAnchor
    Implements IAnchor

    ''' <summary>
    ''' Создаёт объект якоря привязки контента графики для формата OfficeOpenXML.
    ''' </summary>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Создаёт объект якоря привязки контента графики для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="fromCol">Номер начальной колонки</param>
    ''' <param name="fromColOff">Сдвиг начальной колонки. 1 пункт = 12700</param>
    ''' <param name="fromRow">Номер начальной строки</param>
    ''' <param name="fromRowOff">Сдвиг начальной строки. 1 пункт = 12700</param>
    ''' <param name="toCol">Номер конечной колонки</param>
    ''' <param name="toColOff">Сдвиг конечной колонки. 1 пункт = 12700</param>
    ''' <param name="toRow">Номер конечной строки</param>
    ''' <param name="toRowOff">Сдвиг конечной строки. 1 пункт = 12700</param>
    Public Sub New(fromCol As Integer, fromColOff As Integer, fromRow As Integer, fromRowOff As Integer,
            toCol As Integer, toColOff As Integer, toRow As Integer, toRowOff As Integer)
        Me.FromCol = fromCol
        Me.FromColOff = fromColOff
        Me.FromRow = fromRow
        Me.FromRowOff = fromRowOff
        Me.ToCol = toCol
        Me.ToColOff = toColOff
        Me.ToRow = toRow
        Me.ToRowOff = toRowOff
    End Sub

    Friend ReadOnly Property Type As String Implements IAnchor.Type
        Get
            Return "twoCellAnchor"
        End Get
    End Property

    ''' <summary>
    ''' Номер начальной колонки.
    ''' </summary>
    Public Property FromCol As Integer

    ''' <summary>
    ''' Сдвиг начальной колонки. 1 пункт = 12700.
    ''' </summary>
    Public Property FromColOff As Integer

    ''' <summary>
    ''' Номер начальной строки.
    ''' </summary>
    Public Property FromRow As Integer

    ''' <summary>
    ''' Сдвиг начальной строки. 1 пункт = 12700.
    ''' </summary>
    Public Property FromRowOff As Integer

    ''' <summary>
    ''' Номер конечной колонки.
    ''' </summary>
    Public Property ToCol As Integer

    ''' <summary>
    ''' Сдвиг конечной колонки. 1 пункт = 12700.
    ''' </summary>
    Public Property ToColOff As Integer

    ''' <summary>
    ''' Номер конечной строки.
    ''' </summary>
    Public Property ToRow As Integer

    ''' <summary>
    ''' Сдвиг конечной строки. 1 пункт = 12700.
    ''' </summary>
    Public Property ToRowOff As Integer

End Class

