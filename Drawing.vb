Imports OpenXML.Zip
Imports OpenXML.Relationship
Imports OpenXML.ContentType
Imports OpenXML.XML

''' <summary>
''' Объект графики для формата OfficeOpenXML.
''' </summary>
Public Class Drawing

    Private _doc As XDocument
    Private _contens As List(Of IContentDrawing)
    Private _sheet As Sheet
    Private _ct As ContentTypes
    Private _rels As Relationships

    Sub New()
        _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                    {New XElement(NS.xdr + "wsDr",
                                  New XAttribute(XNamespace.Xmlns + "a", NS.a.ToString),
                                  New XAttribute(XNamespace.Xmlns + "r", NS.r.ToString))})
        _contens = New List(Of IContentDrawing)
    End Sub

    ''' <summary>
    ''' Идентификатор связи графики.
    ''' </summary>
    Public Property Id As String

    ''' <summary>
    ''' Возвращает XML файл.
    ''' </summary>
    Public Function GetXML() As XDocument
        Return _doc
    End Function

    ''' <summary>
    ''' Добавляет в Zip-архив файлы графики, контента графики и связи.
    ''' </summary>
    ''' <param name="zip">Zip-архив</param>
    ''' <returns>Возвращает ZIP-архив</returns>
    Public Function AddZip(zip As ZipFile) As ZipFile
        zip.CreateDirectory("drawings", "xl")
        zip.CreateFile("drawing" & _sheet.SheetId & ".xml", _doc, "xl\drawings")
        If Not IsNothing(_rels) Then _rels.AddZip(zip)

        For Each c In _contens
            c.AddZip(zip)
        Next

        Return zip
    End Function

    ''' <summary>
    ''' Добавляет контент графики в XML докумаент.
    ''' </summary>
    ''' <param name="content">Контент графики</param>
    ''' <param name="contentType">Тип контента</param>
    ''' <param name="relationship">Связь</param>
    Private Sub AddContentToXML(content As IContentDrawing, contentType As IDefaultContentType, relationship As IRelationship)
        _ct.Add(contentType)
        Dim rels = relationship
        _rels.Add(rels)
        content.Id = rels.Id
        _doc.Root.Add(New ContentDrawingXML(content))
    End Sub

    ''' <summary>
    ''' Задаёт лист.
    ''' </summary>
    ''' <param name="sheet">Лист</param>
    Public Sub SetSheet(sheet As Sheet)
        _sheet = sheet
        If IsNothing(_ct) Then _ct = _sheet.GetContentTypes
        _ct.Add(New DrawingContentType)
        If IsNothing(_rels) Then _rels = New Relationships("drawing" & _sheet.SheetId & ".xml", "xl\drawings")

        For Each c In _contens
            If c.Type = "pic" Then
                If CType(c, Picture).FormatImg = "png" Then
                    AddContentToXML(c, New PngContentType, New PngRelationship)
                End If
            End If
        Next
    End Sub

    ''' <summary>
    ''' Добавляет контент графики.
    ''' </summary>
    ''' <param name="content">Контент графики</param>
    Public Sub AddContent(content As IContentDrawing)
        If IsNothing(content.Anchor) Then Throw New Exception("Незадан якорь привязки контента графики.")
        _contens.Add(content)
        content.IdContent = _contens.Count

        If Not IsNothing(_sheet) Then
            If IsNothing(_rels) Then _rels = New Relationships("drawing" & _sheet.SheetId & ".xml", "xl\drawings")

            If content.Type = "pic" Then
                If CType(content, Picture).FormatImg = "png" Then
                    AddContentToXML(content, New PngContentType, New PngRelationship)
                End If
            End If
        End If
    End Sub

End Class

