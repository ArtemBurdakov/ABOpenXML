Imports OpenXML.Zip
Imports OpenXML.Style
Imports OpenXML.Relationship
Imports OpenXML.ContentType
Imports OpenXML.XML

''' <summary>
''' Лист Excel для формата OfficeOpenXML.
''' </summary>
Public Class Sheet

    Private _doc As XDocument
    Private _wb As WorkBook
    Private _ct As ContentTypes
    Private _rels As Relationships
    Private _rows As List(Of SheetRow)
    Private _rowN As Integer
    Private _cellN As String
    Private _ss As SharedStrings
    Private _frozenArea As FrozenArea
    Private _cols As List(Of SheetColumn)
    Private _autoFilter As String
    Private _mergeCells As List(Of MergeCell)
    Private _pageMargins As PageMargins
    Private _pageSetup As PageSetup
    Private _drawing As Drawing
    Private _hyperlinks As Hyperlinks

    ''' <summary>
    ''' Создаёт лист Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New()
        _rows = New List(Of SheetRow)
        SummaryBelow = False
    End Sub

    ''' <summary>
    ''' Наименование листа.
    ''' </summary>
    Public Property Name As String

    ''' <summary>
    ''' Порядковый номер листа.
    ''' </summary>
    Public Property SheetId As Integer

    ''' <summary>
    ''' Флаг, указывающий, отображаются ли итоговые строки под своднами данными (по умолчанию false).
    ''' </summary>
    Public Property SummaryBelow As Boolean

    ''' <summary>
    ''' Идентификатор связи.
    ''' </summary>
    Public Property Id As String

    ''' <summary>
    ''' Возвращает XML файл.
    ''' </summary>
    Public Function GetXML() As XDocument
        Return _doc
    End Function

    ''' <summary>
    ''' Возвращает cтроки листа.
    ''' </summary>
    Public Function GetRows() As List(Of SheetRow)
        Return _rows
    End Function

    ''' <summary>
    ''' Добавляет в Zip-архив файлы листов книги и файлы связи.
    ''' </summary>
    ''' <param name="zip">Zip-архив</param>
    ''' <returns>Zip-архив</returns>
    Public Function AddZip(zip As ZipFile) As ZipFile
        CreateXML()
        zip.CreateFile("sheet" & SheetId & ".xml", _doc, "xl\worksheets")
        If Not IsNothing(_rels) Then _rels.AddZip(zip)
        If Not IsNothing(_drawing) Then _drawing.AddZip(zip)
        Return zip
    End Function

    ''' <summary>
    ''' Создаёт из этого объекта XML.
    ''' </summary>
    Private Sub CreateXML()
        _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                    {New XElement(NS.wb + "worksheet",
                                  New XAttribute(XNamespace.Xmlns + "r", NS.r.ToString))})

        AddOutlineParameterToXML()
        AddFrozenAreaToXML()
        AddColumnsWidthToXML()
        ProcessingRows()
        AddRowsToXML()
        AddAutoFilterToXML()
        AddMergeCellsToXML()
        AddPageMarginsToXML()
        AddPageSetupToXML()
        AddDrawingToXML()
        AddHyperlinksToXML()
    End Sub

    ''' <summary>
    ''' Добавляет параметры дерева строк в XML.
    ''' </summary>
    Private Sub AddOutlineParameterToXML()
        If Not SummaryBelow Then
            _doc.Root.Add(New XElement(NS.wb + "sheetPr",
                                           New XElement(NS.wb + "outlinePr",
                                                        New XAttribute("summaryBelow", 0))))
        End If
    End Sub
    ''' <summary>
    ''' Объеденяет ячейки в строки.
    ''' </summary>
    Private Sub ProcessingRows()
        Dim prevRowN = 0
        Dim row = New SheetRow(0)
        Dim result = New List(Of SheetRow)

        For Each g In _rows.OrderBy(Function(x) x.Number).GroupBy(Function(x) x.Number)
            For Each r In g
                If prevRowN <> r.Number Then
                    If prevRowN <> 0 Then
                        row.GetCells.OrderBy(Function(c) c.Number)
                        result.Add(row)
                    End If
                    row = r
                Else
                    row.AddCell(r.CellAt(0))
                End If
                prevRowN = r.Number
            Next
        Next

        row.GetCells.OrderBy(Function(c) c.Number)
        result.Add(row)
        _rows = result
    End Sub

    ''' <summary>
    ''' Добавляет строки в XML документ этого объекта.
    ''' </summary>
    Private Sub AddRowsToXML()
        Dim sheetData = New XElement(NS.wb + "sheetData")
        For Each r In _rows
            Dim row = New SheetRowXML(r)
            For Each c In r.GetCells
                row.Add(New SheetCellXML(c))
            Next
            sheetData.Add(row)
        Next
        _doc.Root.Add(sheetData)
    End Sub

    Private Sub AddColumnsWidthToXML()
        If _cols IsNot Nothing Then
            Dim cols = New XElement(NS.wb + "cols")

            For Each c In _cols
                cols.Add(New SheetColumnXML(c))
            Next

            _doc.Root.Add(cols)
        End If
    End Sub

    Private Sub AddAutoFilterToXML()
        If _autoFilter IsNot Nothing Then
            _doc.Root.Add(New XElement(NS.wb + "autoFilter", New XAttribute("ref", _autoFilter)))
        End If
    End Sub

    Private Sub AddMergeCellsToXML()
        If _mergeCells IsNot Nothing Then
            Dim mergeCells = New XElement(NS.wb + "mergeCells",
                                           New XAttribute("count", _mergeCells.Count))

            For Each mc In _mergeCells
                mergeCells.Add(New MergeCellXML(mc))
            Next

            _doc.Root.Add(mergeCells)
        End If
    End Sub

    Private Sub AddPageMarginsToXML()
        If _pageMargins IsNot Nothing Then
            _doc.Root.Add(_pageMargins.GetXElement())
        End If
    End Sub

    Private Sub AddPageSetupToXML()
        If _pageSetup IsNot Nothing Then
            _doc.Root.Add(_pageSetup.GetXElement())
        End If
    End Sub

    Private Sub AddDrawingToXML()
        If _drawing IsNot Nothing Then
            _doc.Root.Add(New DrawingXML(_drawing))
        End If
    End Sub

    Private Sub AddHyperlinksToXML()
        If _hyperlinks IsNot Nothing Then
            _doc.Root.Add(_hyperlinks.GetXElement())
        End If
    End Sub

    Private Sub AddFrozenAreaToXML()
        If _frozenArea IsNot Nothing Then
            _doc.Root.Add(_frozenArea.GetXElement())
        End If
    End Sub

    ''' <summary>
    ''' Задаёт рабочию книгу.
    ''' </summary>
    ''' <param name="wb">Рабочая книга</param>
    Public Sub SetWorkBook(wb As WorkBook)
        _wb = wb
        _ss = _wb.GetSharedStrings
        _ct = _wb.GetContentTypes
        _ct.Add(New SheetContentType)
    End Sub

    ''' <summary>
    ''' Возвращает список типов контента.
    ''' </summary>
    Public Function GetContentTypes() As ContentTypes
        Return _ct
    End Function

    ''' <summary>
    ''' Задаёт закреплённую область.
    ''' </summary>
    ''' <param name="xSplit">Количество фиксированных столбцов</param>
    ''' <param name="ySplit">Количество фиксированных строк</param>
    Public Sub SetFrozenArea(xSplit As Integer, ySplit As Integer)
        _frozenArea = New FrozenArea()
        _frozenArea.XSplit = xSplit
        _frozenArea.YSplit = ySplit
    End Sub

    ''' <summary>
    ''' Задаёт ширину столбцам.
    ''' </summary>
    ''' <param name="numberColumn">Порядковый номер стобца начиная с 1</param>
    ''' <param name="width">Ширина столбца</param>
    Public Sub SetColumnsWidth(numberColumn As Integer, width As Double)
        SetColumnsWidth(numberColumn, numberColumn, width)
    End Sub

    ''' <summary>
    ''' Задаёт ширину столбцам.
    ''' </summary>
    ''' <param name="startColumn">Порядковый номер стартового стобца начиная с 1</param>
    ''' <param name="endColumn">Порядковый номер последнего стобца начиная с 1</param>
    ''' <param name="width">Ширина столбца</param>
    Public Sub SetColumnsWidth(startColumn As Integer, endColumn As Integer, width As Double)
        If IsNothing(_cols) Then
            _cols = New List(Of SheetColumn)
        End If

        If startColumn > endColumn Then
            Throw New Exception("Порядковый номер стартового стобца выше порядкового номера последнего столбца.")
        End If

        If _cols.Where(Function(x) x.Min <= startColumn And x.Min <= endColumn And x.Max >= startColumn).Count > 0 Then
            Throw New Exception("Некоторым столбцам из диапозона от " & startColumn & " до " & endColumn & " уже задана ширина.")
        End If

        _cols.Add(New SheetColumn(startColumn, endColumn, width))
    End Sub

    ''' <summary>
    ''' Задаёт фильтр для таблицы. Пример: A3:M3
    ''' </summary>
    ''' <param name="ref">Область шапки таблицы</param>
    Public Sub SetAutoFilter(ref As String)
        _autoFilter = ref
        _wb.AddAutoFilterToXML(Name, ref)
    End Sub

    ''' <summary>
    ''' Задаёт объеденённую ячейку.
    ''' </summary>
    ''' <param name="ref">Ссылка на объединёную ячейку в формате A1:B2</param>
    Public Sub SetMergeCell(ref As String)
        If ref.IndexOf(":") = -1 Then
            Throw New Exception("Неверный формат ссылки на объединённую ячейку.")
        Else
            CheckFormatCell(ref.Substring(0, ref.IndexOf(":")))
            Dim r = _rowN
            Dim c = _cellN
            CheckFormatCell(ref.Substring(ref.IndexOf(":") + 1))
            If r > _rowN Then
                Throw New Exception("Неверная последовательность ячеек.")
            Else
                If c.First > _cellN.First Then
                    Throw New Exception("Неверная последовательность ячеек.")
                End If
            End If
        End If

        If IsNothing(_mergeCells) Then
            _mergeCells = New List(Of MergeCell)
        End If

        _mergeCells.Add(New MergeCell(ref))
    End Sub

    ''' <summary>
    ''' Задаёт отступы на печатной странице.
    ''' </summary>
    ''' <param name="footer">Нижний колонтитул</param>
    ''' <param name="header">Верхний колонтитул</param>
    ''' <param name="bottom">Нижний край</param>
    ''' <param name="top">Верхний край</param>
    ''' <param name="right">Правый край</param>
    ''' <param name="left">Левый край</param>
    Public Sub SetPageMargins(footer As Double, header As Double, bottom As Double, top As Double, right As Double, left As Double)
        _pageMargins = New PageMargins() With {
                .Bottom = bottom,
                .Footer = footer,
                .Header = header,
                .Left = left,
                .Right = right,
                .Top = top
            }
    End Sub

    ''' <summary>
    ''' Возвращает элемент находящийся перед PageSetup.
    ''' </summary>
    Private Function SetElementBeforePageSetup() As XElement
        Dim el As XElement
        If Not IsNothing(_doc.Root.Element(NS.wb + "pageMargins")) Then
            el = _doc.Root.Element(NS.wb + "pageMargins")
        ElseIf Not IsNothing(_mergeCells) Then
            el = _doc.Root.Element(NS.wb + "mergeCells")
        Else
            el = _doc.Root.Element(NS.wb + "sheetData")
        End If
        Return el
    End Function

    ''' <summary>
    ''' Задаёт настройки для печатной страницы.
    ''' </summary>
    ''' <param name="orientation">Ориентация страницы</param>
    Public Sub SetPageSetup(orientation As String)
        SetPageSetup(orientation, Nothing)
    End Sub

    ''' <summary>
    ''' Задаёт настройки для печатной страницы.
    ''' </summary>
    ''' <param name="scale">Масштаб страницы</param>
    Public Sub SetPageSetup(scale As Integer)
        SetPageSetup(Nothing, scale)
    End Sub

    ''' <summary>
    ''' Задаёт настройки для печатной страницы.
    ''' </summary>
    ''' <param name="orientation">Ориентация страницы</param>
    ''' <param name="scale">Масштаб страницы</param>
    Public Sub SetPageSetup(orientation As String, scale As Integer)
        _pageSetup = New PageSetup() With {
                .Orientation = orientation,
                .Scale = scale
            }
    End Sub

    ''' <summary>
    ''' Задаёт графику листу.
    ''' </summary>
    ''' <param name="drawing">Графика</param>
    Public Sub SetDrawing(drawing As Drawing)
        _drawing = drawing
        _drawing.SetSheet(Me)
        CreateRelationshipsIfNothing()
        Dim rels = New DrawingRelationship
        _rels.Add(rels)
        _drawing.Id = rels.Id
    End Sub

    ''' <summary>
    ''' Проверяет формат и определят номер строки.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    Private Sub CheckFormatCell(number As String)
        _rowN = Nothing
        _cellN = Nothing
        For i = 0 To number.Count - 1
            If IsNumeric(number.Chars(i)) Then
                If i = 0 Then Throw New Exception("Неверный формат номера ячейки " & number & ".")
                _rowN = number.Substring(i)
                _cellN = number
                Exit For
            End If
        Next
        If IsNothing(_rowN) Then Throw New Exception("Неверный формат номера ячейки " & number & ".")
    End Sub

    ''' <summary>
    ''' Добавляет стиль в список стилей.
    ''' </summary>
    ''' <param name="style">Стиль ячейки</param>
    Private Function AddStyle(style As IStyleCell) As IStyleCell
        If Not IsNothing(_wb) Then
            Return _wb.Styles.Add(style)
        Else
            Throw New Exception("К листу " & Name & " не привязанна книга.")
        End If

        Return Nothing
    End Function

    ''' <summary>
    ''' Возвращает строку с указанным номером.
    ''' </summary>
    ''' <param name="number">Номер строки</param>
    Public Function Row(number As Integer) As SheetRow
        Return _rows.Where(Function(x) x.Number = number).First
    End Function

    ''' <summary>
    ''' Возвращает последнюю строку.
    ''' </summary>
    Public Function LastRow() As SheetRow
        Return _rows.Last
    End Function

    ''' <summary>
    ''' Добавляет строку в конец колекции листа.
    ''' </summary>
    ''' <param name="row">Строка</param>
    Public Sub AddRow(row As SheetRow)
        For Each cell In row.GetCells
            cell.Style = AddStyle(cell.Style)
            _rows.Add(New SheetRow(row))
            _rows.Last.AddCell(cell)
        Next
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="style">Стиль ячейки</param>
    Public Sub AddCell(number As String, style As IStyleCell)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, "", style:=AddStyle(style)))
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="value">Содержимое ячейки</param>
    Public Sub AddCell(number As String, value As Object)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, value))
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="value">Содержимое ячейки</param>
    ''' <param name="style">Стиль ячейки</param>
    Public Sub AddCell(number As String, value As Object, style As IStyleCell)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, value, style:=AddStyle(style)))
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="value">Содержимое ячейки</param>
    ''' <param name="height">Высота строки</param>
    Public Sub AddCell(number As String, value As Object, height As Double)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, value, height))
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="value">Содержимое ячейки</param>
    ''' <param name="height">Высота строки</param>
    ''' <param name="style">Стиль ячейки</param>
    Public Sub AddCell(number As String, value As Object, height As Double, style As IStyleCell)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, value, height, AddStyle(style)))
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="value">Содержимое ячейки</param>
    ''' <param name="style">Стиль ячейки</param>
    Public Sub AddCell(number As String, value As Date, style As IStyleCell)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, value, style:=AddStyle(style)))
    End Sub

    ''' <summary>
    ''' Добавляет ячейку в лист.
    ''' </summary>
    ''' <param name="number">Номер ячейки в формате "A1"</param>
    ''' <param name="value">Содержимое ячейки</param>
    ''' <param name="height">Высота строки</param>
    ''' <param name="style">Стиль ячейки</param>
    Public Sub AddCell(number As String, value As Date, height As Double, style As IStyleCell)
        CheckFormatCell(number)
        _rows.Add(New SheetRow(_rowN, _cellN, value, height, AddStyle(style)))
    End Sub

    ''' <summary>
    ''' Добавляет гиперссылку.
    ''' </summary>
    ''' <param name="cell">Ячейка для гиперссылки</param>
    ''' <param name="path">Путь гиперссылки</param>
    Public Sub AddHyperlink(cell As SheetCell, path As String)
        ChangeStyleHyperlink(cell)
        CreateRelationshipsIfNothing()
        Dim relsh = New HyperlinkRelationship
        relsh.Target = path
        _rels.Add(relsh)
        If _hyperlinks Is Nothing Then
            _hyperlinks = New Hyperlinks
        End If
        _hyperlinks.Add(New Hyperlink With {.Id = relsh.Id, .NumberCell = cell.Number})
    End Sub

    ''' <summary>
    ''' Возвращает ячейку с указаным номером или вызывает исключение, если она не задана.
    ''' </summary>
    ''' <param name="numberCell">Номер ячейки</param>
    Private Function GetCell(numberCell As String) As SheetCell
        Dim row = _rows.Where(Function(x) x.CellAt(0).Number = numberCell).FirstOrDefault
        If row IsNot Nothing Then
            Return row.CellAt(0)
        End If
        Throw New Exception("Ячейка с номером " & numberCell & " в листе " & Name & " не задана.")
    End Function

    ''' <summary>
    ''' Изменяет стиль ячейки для гиперссылки.
    ''' </summary>
    ''' <param name="cell">Ячейка с гиперссылкой</param>
    Private Sub ChangeStyleHyperlink(cell As SheetCell)
        Dim style = New StyleCell(cell.Style)
        style.Font.Underline = True
        style.Font.Color = New Color With {.Theme = 10}
        AddStyle(style)
        cell.Style = style
    End Sub

    ''' <summary>
    ''' Создаёт список связей для этого объекта, если он ещё не создан.
    ''' </summary>
    Private Sub CreateRelationshipsIfNothing()
        If IsNothing(_rels) Then
            _rels = New Relationships("sheet" & SheetId & ".xml", "xl\worksheets")
        End If
    End Sub

End Class

