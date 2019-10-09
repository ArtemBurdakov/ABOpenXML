''' <summary>
''' Объединённая ячейка листа Excel для формата OfficeOpenXML.
''' </summary>
Class MergeCell

    ''' <summary>
    ''' Создаёт объект объединённой ячейки листа Excel для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="ref">Ссылка на объединёную ячейку в формате A1:B2</param>
    Sub New(ref As String)
        Me.Ref = ref
    End Sub

    ''' <summary>
    ''' Ссылка на объединёную ячейку в формате A1:B2.
    ''' </summary>
    Public Property Ref As String

End Class

