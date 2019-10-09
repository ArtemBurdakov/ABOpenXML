Imports System.IO
Imports OpenXML.Zip

''' <summary>
''' Рисунок для формата OfficeOpenXML.
''' </summary>
Public Class Picture
    Implements IContentDrawing

    Private _pathImg As String
    Private _dtaFile As Byte()

    ''' <summary>
    ''' Создаёт объект рисунка для формата OfficeOpenXML.
    ''' </summary>
    Public Sub New()

    End Sub

    ''' <summary>
    ''' Создаёт объект рисунка для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="pathImg">Путь к изображению</param>
    ''' <param name="name">Наименование</param>
    Public Sub New(pathImg As String, name As String)
        _pathImg = pathImg
        FormatImg = pathImg.Substring(pathImg.LastIndexOf("."c) + 1)
        Me.Name = name
    End Sub

    ''' <summary>
    ''' Создаёт объект рисунка для формата OfficeOpenXML.
    ''' </summary>
    ''' <param name="dataFile">Данные изображения в виде массива байтов</param>
    ''' <param name="formatImg">Путь к изображению</param>
    ''' <param name="name">Наименование</param>
    Public Sub New(dataFile As Byte(), formatImg As String, name As String)
        _dtaFile = dataFile
        Me.FormatImg = formatImg
        Me.Name = name
    End Sub

    ''' <summary>
    ''' Идентификатор связи рисунка.
    ''' </summary>
    Friend Property Id As String Implements IContentDrawing.Id

    ''' <summary>
    ''' Идентификатор контента графики.
    ''' </summary>
    Friend Property IdContent As Integer Implements IContentDrawing.IdContent

    ''' <summary>
    ''' Наименование картинки.
    ''' </summary>
    Public Property Name As String Implements IContentDrawing.Name

    ''' <summary>
    ''' Вид контента графики.
    ''' </summary>
    Friend ReadOnly Property Type As String Implements IContentDrawing.Type
        Get
            Return "pic"
        End Get
    End Property

    ''' <summary>
    ''' Якорь привязки контента графики.
    ''' </summary>
    Public Property Anchor As IAnchor Implements IContentDrawing.Anchor

    ''' <summary>
    ''' Формат изображения.
    ''' </summary>
    Public Property FormatImg As String

    ''' <summary>
    ''' Добавляет в Zip-архив файл изображения.
    ''' </summary>
    ''' <param name="zip">Zip-архив</param>
    ''' <returns>Возвращает ZIP-архив</returns>
    Friend Function AddZip(zip As ZipFile) As ZipFile Implements IContentDrawing.AddZip
        zip.CreateDirectory("media", "xl")
        If Not IsNothing(_pathImg) Then
            zip.CreateFile("image" & IdContent & "." & FormatImg, File.ReadAllBytes(_pathImg), "xl\media")
        ElseIf Not IsNothing(_dtaFile)
            zip.CreateFile("image" & IdContent & "." & FormatImg, _dtaFile, "xl\media")
        End If
        Return zip
    End Function

End Class

