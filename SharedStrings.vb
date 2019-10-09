Imports OpenXML.Zip
Imports OpenXML.XML

''' <summary>
''' Общие строки книги Excel для формата OfficeOpenXML.
''' </summary>
Public Class SharedStrings

    Private _doc As XDocument
    Private _listRows As List(Of IEnumerable(Of SheetRow))
    Private _strings As List(Of SharedString)

    ''' <summary>
    ''' Создаёт объект общие строки книги Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New()
        _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                    {New XElement(NS.wb + "sst",
                                  New XAttribute("count", 0),
                                  New XAttribute("uniqueCount", 0))})
        _strings = New List(Of SharedString)
        _listRows = New List(Of IEnumerable(Of SheetRow))
        Count = 0
        UniqueCount = 0
    End Sub

    ''' <summary>
    ''' Количество ссылок на строки.
    ''' </summary>
    Public Property Count As Integer

    ''' <summary>
    ''' Количество уникальных строк.
    ''' </summary>
    Public Property UniqueCount As Integer

    ''' <summary>
    ''' Возвращает XML файл.
    ''' </summary>
    Public Function GetXML() As XDocument
        Return _doc
    End Function

    ''' <summary>
    ''' Обрабатывает данные для нахождения общих строк.
    ''' </summary>
    Private Sub ProcessingSharedStrings()
        Dim prevString As String = Nothing
        Dim uRows = _listRows.ElementAt(0)

        For i = 1 To _listRows.Count - 1
            uRows = uRows.Union(_listRows.ElementAt(i))
        Next

        Dim row As SheetRow
        For Each gRow In uRows.Where(Function(x) x.CellAt(0).Type = "s").
                OrderBy(Function(x) x.GetCells.First.Value).GroupBy(Function(x) x.GetCells.First.Value)
            For Each row In gRow
                If prevString IsNot Nothing AndAlso prevString = row.CellAt(0).Value Then
                    row.CellAt(0).Value = _strings.Last.Number
                Else
                    prevString = row.CellAt(0).Value
                    _strings.Add(New SharedString(UniqueCount, row.CellAt(0).Value))
                    _doc.Root.Add(New SharedStringXML(_strings.Last))
                    row.CellAt(0).Value = UniqueCount
                    UniqueCount += 1
                End If
                Count += 1
            Next
        Next
    End Sub

    ''' <summary>
    ''' Добавляет в Zip-архив файл общих строк.
    ''' </summary>
    ''' <param name="zip">Zip-архив</param>
    ''' <returns>Zip-архив</returns>
    Public Function AddZip(zip As ZipFile) As ZipFile
        ProcessingSharedStrings()
        _doc.Root.Attribute("count").Value = Count
        _doc.Root.Attribute("uniqueCount").Value = UniqueCount
        zip.CreateFile("sharedStrings.xml", _doc, "xl")
        Return zip
    End Function

    ''' <summary>
    ''' Добавляет строки листа Excel для последующей обработки.
    ''' </summary>
    ''' <param name="rows">Строки листа</param>
    Public Sub AddRows(rows As List(Of SheetRow))
        _listRows.Add(rows)
    End Sub

End Class

