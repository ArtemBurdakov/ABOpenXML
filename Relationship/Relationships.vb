Imports OpenXML.Zip
Imports OpenXML.XML

Namespace Relationship

    ''' <summary>
    ''' Описатель связей для формата OfficeOpenXML.
    ''' </summary>
    Class Relationships

        Private _doc As XDocument

        ''' <summary>
        ''' Создаёт файл описатель связей для формата OfficeOpenXML.
        ''' </summary>
        ''' <param name="name">Имя файла</param>
        ''' <param name="path">Путь к файлу (без имени)</param>
        Sub New(name As String, path As String)
            Me.Name = name
            Me.Path = path
            _doc = New XDocument(New XDeclaration("1.0", "UTF-8", "yes"),
                                    {New XElement(NS.rs + "Relationships")})
        End Sub

        ''' <summary>
        ''' Имя файла
        ''' </summary>
        Public Property Name As String

        ''' <summary>
        ''' Путь к файлу (без имени)
        ''' </summary>
        Public Property Path As String

        ''' <summary>
        ''' Возвращает XML файл.
        ''' </summary>
        Public Function GetXML() As XDocument
            Return _doc
        End Function

        ''' <summary>
        ''' Добавляет в Zip-архив директорию и файл связи.
        ''' </summary>
        ''' <param name="zip">Zip-архив</param>
        ''' <returns>Возвращает ZIP-архив</returns>
        Public Function AddZip(zip As ZipFile) As ZipFile
            Dim _path = If(Path = "", "", If(Path.Last = "\", Path, Path & "\"))
            zip.CreateDirectory("_rels", _path)
            zip.CreateFile(Name & ".rels", _doc, _path & "_rels")
            Return zip
        End Function

        ''' <summary>
        ''' Добавляет связь.
        ''' </summary>
        ''' <param name="rel">Связь</param>
        Public Sub Add(rel As IRelationship)
            Dim rels = _doc.Root.Elements(NS.rs + "Relationship") _
                           .Where(Function(x) x.Attribute("Type").Value = rel.Type)
            rel.Id = "rId" & _doc.Root.Elements(NS.rs + "Relationship").Count + 1

            If (rels.Count = 0 AndAlso rel.Unique) OrElse (Not rel.Unique AndAlso rel.TargetMode = "External") Then
                _doc.Root.Add(New RelationshipXML(rel))
            ElseIf Not rel.Unique AndAlso rel.TargetMode <> "External" Then
                Dim indexDot = rel.Target.LastIndexOf("."c)
                rel.Target = rel.Target.Substring(0, indexDot) & (rels.Count + 1) & rel.Target.Substring(indexDot)
                _doc.Root.Add(New RelationshipXML(rel))
            End If
        End Sub

    End Class

End Namespace

