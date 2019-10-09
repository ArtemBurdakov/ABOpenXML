Imports System.IO
Imports OpenXML.Zip
Imports OpenXML.Relationship
Imports OpenXML.ContentType

''' <summary>
''' Формат документа OfficeOpenXML.
''' </summary>
Public Class OfficeOpenXML

    Private _ct As ContentTypes
    Private _rels As Relationships
    Private _workBook As WorkBook
    Private _dataBytes As Byte()

    ''' <summary>
    ''' Создаёт объект документа формата OfficeOpenXML.
    ''' </summary>
    Public Sub New()
        _ct = New ContentTypes
        _rels = New Relationships("", "")
    End Sub

    ''' <summary>
    ''' Возвращает книгу Excel.
    ''' </summary>
    Public ReadOnly Property WorkBook() As WorkBook
        Get
            If IsNothing(_workBook) Then
                _workBook = New WorkBook()
                _rels.Add(New WorkBookRelationship)
                _workBook.SetContentTypes(_ct)
            End If
            Return _workBook
        End Get
    End Property

    ''' <summary>
    ''' Выполнить архивацию документа.
    ''' </summary>
    Private Sub Zip()
        Dim zip = New ZipFile
        _ct.AddZip(zip)
        _rels.AddZip(zip)
        _workBook.AddZip(zip)
        _dataBytes = zip.Archive
    End Sub

    ''' <summary>
    ''' Сохраняет файл по указаннуму пути.
    ''' </summary>
    ''' <param name="path">Путь к файлу</param>
    Public Sub Save(path As String)
        Zip()
        If IsNothing(_dataBytes) Then Throw New Exception("Не выполнена архивация документа.")
        File.WriteAllBytes(path, _dataBytes)
        Process.Start(path)
    End Sub

    ''' <summary>
    ''' Открывает файл по указаннуму пути.
    ''' </summary>
    ''' <param name="path">Путь к файлу</param>
    Public Sub Open(path As String)
        Save(path)
        Process.Start(path)
    End Sub

    ''' <summary>
    ''' Возвращает массив байтов Excel документа.
    ''' </summary>
    Public Function GetByte() As Byte()
        Zip()
        If IsNothing(_dataBytes) Then Throw New Exception("Не выполнена архивация документа.")
        Return _dataBytes
    End Function

End Class

