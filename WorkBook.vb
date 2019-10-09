Imports OpenXML.Zip
Imports OpenXML.Relationship
Imports OpenXML.ContentType
Imports OpenXML.XML

''' <summary>
''' Книга Excel для формата OfficeOpenXML.
''' </summary>
Public Class WorkBook

    Private _doc As XDocument
    Private _ct As ContentTypes
    Private _rels As Relationships
    Private _sheets As IList(Of Sheet)
    Private _ss As SharedStrings
    Private _styles As StyleSheet

    ''' <summary>
    ''' Создаёт книгу Excel для формата OfficeOpenXML.
    ''' </summary>
    Sub New()
        _sheets = New List(Of Sheet)
        _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                    {New XElement(NS.wb + "workbook",
                                  New XAttribute(XNamespace.Xmlns + "r", NS.r.ToString),
                                  New XElement(NS.wb + "sheets"))})
        _ss = New SharedStrings()
        _rels = New Relationships("workbook.xml", "xl")
        _rels.Add(New SharedStringsRelationship)
    End Sub

    Public ReadOnly Property Styles As StyleSheet
        Get
            If IsNothing(_styles) Then
                _styles = New StyleSheet
                If Not IsNothing(_ct) Then _ct.Add(New StylesContentType)
                _rels.Add(New StylesRelationship)
            End If
            Return _styles
        End Get
    End Property

    ''' <summary>
    ''' Возвращает XML файл.
    ''' </summary>
    Public Function GetXML() As XDocument
        Return _doc
    End Function

    ''' <summary>
    ''' Добавляет в Zip-архив директории xl, xl/worksheets, файлы workbook.xml, sharedStrings.xml и листов книги и файлы связи.
    ''' </summary>
    ''' <param name="zip">Zip-архив</param>
    ''' <returns>Возвращает ZIP-архив</returns>
    Public Function AddZip(zip As ZipFile) As ZipFile
        zip.CreateDirectory("xl")
        zip.CreateFile("workbook.xml", _doc, "xl")
        _rels.AddZip(zip)

        If _sheets.Count > 0 Then
            zip.CreateDirectory("worksheets", "xl")

            For Each s In _sheets
                _ss.AddRows(s.GetRows)
            Next

            _ss.AddZip(zip)
            If Not IsNothing(_styles) Then _styles.AddZip(zip)

            For Each s In _sheets
                s.AddZip(zip)
            Next
        End If

        Return zip
    End Function

    Friend Sub AddAutoFilterToXML(nameSheet As String, autoFilter As String)
        _doc.Root.Add(New XElement(NS.wb + "definedNames",
                New XElement(NS.wb + "definedName",
                             nameSheet & "!$" & autoFilter.Substring(0, 1) & "$" & autoFilter.Substring(1, 1) & ":" &
                             autoFilter.Substring(3, 1) & "$" & autoFilter.Substring(4, 1),
                             New XAttribute("name", "_xlnm._FilterDatabase"),
                             New XAttribute("localSheetId", 0), New XAttribute("hidden", 1))))
    End Sub

    ''' <summary>
    ''' Задаёт список типов контента этому объекту.
    ''' </summary>
    ''' <param name="ct">Список типов контента</param>
    Public Sub SetContentTypes(ct As ContentTypes)
        _ct = ct
        _ct.Add(New RelsContentType)
        _ct.Add(New XmlContentType)
        _ct.Add(New WorkBookContentType)
        _ct.Add(New SharedStringsContentType)
        If Not IsNothing(_styles) Then _ct.Add(New StylesContentType)
    End Sub

    ''' <summary>
    ''' Возвращает список типов контента.
    ''' </summary>
    Public Function GetContentTypes() As ContentTypes
        Return _ct
    End Function

    ''' <summary>
    ''' Возвращает общие строки.
    ''' </summary>
    Public Function GetSharedStrings() As SharedStrings
        Return _ss
    End Function

    ''' <summary>
    '''  Добавляет лист в книгу.
    ''' </summary>
    ''' <param name="name">Наименование листа</param>
    Public Function AddSheet(name As String) As Sheet
        If IsNothing(_ct) Then Throw New Exception("Не задан список типов контента.")
        Dim sheetRelationship = New SheetRelationship
        _rels.Add(sheetRelationship)
        Dim sheet = New Sheet
        _sheets.Add(sheet)
        sheet.Name = name
        sheet.SheetId = _sheets.Count
        sheet.Id = sheetRelationship.Id
        sheet.SetWorkBook(Me)
        _doc.Root.Element(NS.wb + "sheets").Add(New SheetXML(sheet))
        Return sheet
    End Function

    ''' <summary>
    ''' Возвращает лист по указанному индексу из книги.
    ''' </summary>
    ''' <param name="index">Индекс</param>
    Public Function SheetAt(index As Integer) As Sheet
        Return _sheets.ElementAt(index)
    End Function

End Class

