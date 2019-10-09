Imports OpenXML.Zip
Imports OpenXML.Style
Imports OpenXML.XML

''' <summary>
''' Список стилей в Excel для формата OfficeOpenXML.
''' </summary>
Public Class StyleSheet

    Private _doc As XDocument
    Private _cellStyles As List(Of IStyleCell)
    Private _cellXfs As List(Of IStyleCell)
    Private _alignments As List(Of AlignmentStyle)
    Private _fonts As List(Of FontStyle)
    Private _fills As List(Of FillStyle)
    Private _borders As List(Of BorderStyle)
    Private _numFmts As List(Of NumFmtStyle)

    ''' <summary>
    ''' Создаёт объект список стилей в Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New()
        _cellStyles = New List(Of IStyleCell)
        _cellXfs = New List(Of IStyleCell)
        CellStylesCount = 0
        CellXfsCount = 0

        _alignments = New List(Of AlignmentStyle) From {New AlignmentStyle}
        _alignments.Last.Id = 0

        _fonts = New List(Of FontStyle) From {New FontStyle}
        _fonts.First.Id = 0

        _fills = New List(Of FillStyle) From {New FillStyle, New FillStyle(PatternTypeStyle.Gray125)}
        _fills.First.Id = 0
        _fills.Last.Id = 1

        _borders = New List(Of BorderStyle) From {New BorderStyle}
        _borders.First.Id = 0

        _numFmts = New List(Of NumFmtStyle)

        Dim cellStyle = New StyleCell(0, CellStylesCount, Nothing, "Основной")
        cellStyle.Alignment = _alignments.Last
        cellStyle.Font = _fonts.First
        cellStyle.Fill = _fills.First
        cellStyle.Border = _borders.First
        _cellStyles.Add(cellStyle)
        CellStylesCount += 1

        _cellXfs.Add(New StyleCell(1, CellXfsCount, 0, Nothing))
        _cellXfs.First.Alignment = _alignments.Last
        _cellXfs.First.Font = _fonts.Last
        _cellXfs.First.Fill = _fills.First
        _cellXfs.First.Border = _borders.First
        CellXfsCount += 1

        _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                    {New XElement(NS.wb + "styleSheet",
                                  New XElement(NS.wb + "fonts",
                                               New XAttribute("count", _fonts.Count)),
                                  New XElement(NS.wb + "fills",
                                               New XAttribute("count", _fills.Count)),
                                  New XElement(NS.wb + "borders",
                                               New XAttribute("count", _borders.Count)),
                                  New XElement(NS.wb + "cellStyleXfs",
                                               New XAttribute("count", CellStylesCount),
                                               New XfXML(cellStyle)),
                                  New XElement(NS.wb + "cellXfs",
                                               New XAttribute("count", CellXfsCount)),
                                  New XElement(NS.wb + "cellStyles",
                                               New XAttribute("count", CellStylesCount),
                                               New CellStyleXML(cellStyle)))})
    End Sub

    ''' <summary>
    ''' Количество основных стилей ячейки.
    ''' </summary>
    Public Property CellStylesCount As Integer

    ''' <summary>
    ''' Количество прямых стилей ячейки.
    ''' </summary>
    Public Property CellXfsCount As Integer

    ''' <summary>
    ''' Возвращает XML файл.
    ''' </summary>
    Public Function GetXML() As XDocument
        Return _doc
    End Function

    ''' <summary>
    ''' Обработка шрифтов.
    ''' </summary>
    Private Sub ProcessingFonts()
        Dim fonts = _doc.Root.Element(NS.wb + "fonts")

        For Each f In _fonts
            fonts.Add(New FontStyleXML(f))
        Next

        fonts.Attribute("count").Value = _fonts.Count
    End Sub

    ''' <summary>
    ''' Обработка заливки.
    ''' </summary>
    Private Sub ProcessingFills()
        Dim fills = _doc.Root.Element(NS.wb + "fills")

        For Each f In _fills
            fills.Add(New FillStyleXML(f))
        Next

        fills.Attribute("count").Value = _fills.Count
    End Sub

    ''' <summary>
    ''' Обработка границ.
    ''' </summary>
    Private Sub ProcessingBorders()
        Dim borders = _doc.Root.Element(NS.wb + "borders")

        For Each b In _borders
            borders.Add(New BorderStyleXML(b))
        Next

        borders.Attribute("count").Value = _borders.Count
    End Sub

    ''' <summary>
    ''' Обработка форматирования цифр.
    ''' </summary>
    Private Sub ProcessingNumFmts()
        If _numFmts.Count > 0 Then
            _doc.Root.AddFirst(New XElement(NS.wb + "numFmts",
                                                New XAttribute("count", 0)))

            Dim numFmts = _doc.Root.Element(NS.wb + "numFmts")

            For Each n In _numFmts
                numFmts.Add(New NumFmtStyleXML(n))
            Next

            numFmts.Attribute("count").Value = _numFmts.Count
        End If
    End Sub

    ''' <summary>
    ''' Обработка прямых стилей.
    ''' </summary>
    Private Sub ProcessingXfs()
        Dim cellXfs = _doc.Root.Element(NS.wb + "cellXfs")

        For Each c In _cellXfs
            cellXfs.Add(New XfXML(c))
        Next

        CellXfsCount = _cellXfs.Count
        cellXfs.Attribute("count").Value = CellXfsCount
    End Sub

    ''' <summary>
    ''' Добавляет в Zip-архив файл styles.xml.
    ''' </summary>
    ''' <param name="zip">Zip-архив</param>
    ''' <returns>Возвращает ZIP-архив</returns>
    Public Function AddZip(zip As ZipFile) As ZipFile
        ProcessingFonts()
        ProcessingFills()
        ProcessingBorders()
        ProcessingNumFmts()
        ProcessingXfs()
        zip.CreateFile("styles.xml", _doc, "xl")
        Return zip
    End Function

    ''' <summary>
    ''' Добавляет стиль.
    ''' </summary>
    ''' <param name="style">Стиль</param>
    ''' <return>Возвращает добавленный стиль</return>
    Public Function Add(style As IStyleCell) As IStyleCell
        If style.Id = 0 Then
            If IsNothing(style.Alignment) Then
                style.Alignment = _alignments.First
            Else
                Dim alignmentFind = _alignments.Where(Function(x) x = style.Alignment)
                If alignmentFind.Count > 0 Then
                    style.Alignment = alignmentFind.First
                Else
                    style.Alignment.Id = _alignments.Count
                    _alignments.Add(style.Alignment)
                End If
            End If

            If IsNothing(style.Font) Then
                style.Font = _fonts.First
            Else
                Dim fontFind = _fonts.Where(Function(x) x = style.Font)
                If fontFind.Count > 0 Then
                    style.Font = fontFind.First
                Else
                    style.Font.Id = _fonts.Count
                    _fonts.Add(style.Font)
                End If
            End If

            If IsNothing(style.Fill) Then
                style.Fill = _fills.First
            Else
                Dim fillFind = _fills.Where(Function(x) x = style.Fill)
                If fillFind.Count > 0 Then
                    style.Fill = fillFind.First
                Else
                    style.Fill.Id = _fills.Count
                    _fills.Add(style.Fill)
                End If
            End If

            If IsNothing(style.Border) Then
                style.Border = _borders.First
            Else
                Dim borderFind = _borders.Where(Function(x) x = style.Border)
                If borderFind.Count > 0 Then
                    style.Border = borderFind.First
                Else
                    style.Border.Id = _borders.Count
                    _borders.Add(style.Border)
                End If
            End If

            If Not IsNothing(style.NumFmt) Then
                Dim numFmtFind = _numFmts.Where(Function(x) x = style.NumFmt)
                If numFmtFind.Count > 0 Then
                    style.NumFmt = numFmtFind.First
                Else
                    style.NumFmt.Id = _numFmts.Count + 165
                    _numFmts.Add(style.NumFmt)
                End If
            End If

            Dim cellXfFind = _cellXfs.Where(Function(x) x.Alignment.Id = style.Alignment.Id And x.Font.Id = style.Font.Id And x.Fill.Id = style.Fill.Id And
                                                            x.Border.Id = style.Border.Id And If(IsNothing(x.NumFmt), 0, x.NumFmt.Id) = If(IsNothing(style.NumFmt), 0, style.NumFmt.Id))
            If cellXfFind.Count > 0 Then
                style = cellXfFind.First
            Else
                style.Id = _cellXfs.Count
                _cellXfs.Add(style)
                CellXfsCount = _cellXfs.Count
            End If
        End If

        Return style
    End Function

    ''' <summary>
    ''' Возвращает первый найденный по имени основной стиль.
    ''' </summary>
    ''' <param name="name">Наименование стиля</param>
    Public Function Item(name As String) As IStyleCell
        Return _cellXfs.Where(Function(x) x.Name = name).First
    End Function

End Class

