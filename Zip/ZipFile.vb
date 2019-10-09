Imports System.Text
Imports System.IO
Imports System.Xml

Namespace Zip

    ''' <summary>
    ''' Создаёт ZIP-архив
    ''' </summary>
    Public Class ZipFile

        Private _filesZip As New List(Of FileZip)

        ''' <summary>
        ''' Добавляет файл в архив.
        ''' </summary>
        ''' <param name="pathFile">Путь к файлу</param>
        ''' <param name="pathDirectory">Путь к файлу от корня архива</param>
        Public Sub AddFile(pathFile As String, Optional pathDirectory As String = Nothing)
            CreateFile(pathFile.Substring(pathFile.LastIndexOf("\") + 1), File.ReadAllBytes(pathFile), pathDirectory)
        End Sub

        ''' <summary>
        ''' Добавляет директорию в архив
        ''' </summary>
        ''' <param name="pathDirectory">Путь директории</param>
        ''' <param name="pathRootDirectory">Путь к директории от корня архива</param>
        Public Sub AddDirectory(pathDirectory As String, Optional pathRootDirectory As String = "")
            Dim path = If(pathDirectory.Last = "\", pathDirectory, pathDirectory & "\")
            Dim name = path.Substring(path.Remove(path.Length - 1).LastIndexOf("\") + 1)
            name = name.Remove(name.Length - 1)
            Dim nameDirectories = Directory.GetDirectories(path)

            CreateDirectory(name, pathRootDirectory)
            For Each d In nameDirectories
                AddDirectory(d, If(pathRootDirectory = "", "", If(pathRootDirectory.Last = "\", pathRootDirectory, pathRootDirectory & "\")) & name & "\")
            Next

            Dim nameFiles = Directory.GetFiles(path)
            For Each f In nameFiles
                AddFile(f, If(pathRootDirectory = "", "", If(pathRootDirectory.Last = "\", pathRootDirectory, pathRootDirectory & "\")) & name & "\")
            Next
        End Sub

        ''' <summary>
        ''' Создаёт директорию в архиве.
        ''' </summary>
        ''' <param name="nameDirectory">Наименование директории</param>
        ''' <param name="pathDirectory">Путь к директории от корня архива</param>
        Public Sub CreateDirectory(nameDirectory As String, Optional pathDirectory As String = Nothing)
            Dim flag As Boolean = False
            If IsNothing(pathDirectory) Or pathDirectory = "" Then
                _filesZip.Add(New FileZip(nameDirectory & "/"))
                flag = True
            Else
                Dim path = If(pathDirectory.Last = "\", pathDirectory, pathDirectory & "\")
                path = path.Replace("\", "/")
                For Each f In _filesZip
                    If f.Name = path Then
                        _filesZip.Add(New FileZip(path & nameDirectory & "/"))
                        flag = True
                        Exit For
                    End If
                Next
            End If

            If Not flag Then
                Throw New Exception("Директории " & pathDirectory.Replace("\", "/") & " нет в ZIP-архиве.")
            End If
        End Sub

        ''' <summary>
        ''' Создаёт файл в архиве.
        ''' </summary>
        ''' <param name="nameFile">Наименование файла</param>
        ''' <param name="dataFile">Данные файла</param>
        ''' <param name="pathDirectory">Путь к файлу от корня архива</param>
        Public Sub CreateFile(nameFile As String, dataFile As Byte(), Optional pathDirectory As String = Nothing)
            Dim flag As Boolean = False
            If IsNothing(pathDirectory) Or pathDirectory = "" Then
                _filesZip.Add(New FileZip(nameFile, dataFile))
                flag = True
            Else
                Dim path = If(pathDirectory.Last = "\", pathDirectory, pathDirectory & "\")
                path = path.Replace("\", "/")
                For Each f In _filesZip
                    If f.Name = path Then
                        _filesZip.Add(New FileZip(path & nameFile, dataFile))
                        flag = True
                        Exit For
                    End If
                Next
            End If

            If Not flag Then
                Throw New Exception("Директории " & pathDirectory.Replace("\", "/") & " нет в ZIP-архиве.")
            End If
        End Sub

        ''' <summary>
        ''' Создаёт файл в архиве.
        ''' </summary>
        ''' <param name="nameFile">Наименование файла</param>
        ''' <param name="xDoc">Xml документ</param>
        ''' <param name="pathDirectory">Путь к файлу от корня архива</param>
        Public Sub CreateFile(nameFile As String, xDoc As XDocument, Optional pathDirectory As String = Nothing)
            Dim settings = New XmlWriterSettings With {.OmitXmlDeclaration = True, .Encoding = Encoding.UTF8}
            Using memoryStream = New MemoryStream()
                Using xWriter = XmlWriter.Create(memoryStream, settings)
                    xDoc.WriteTo(xWriter)
                    xWriter.Flush()
                    CreateFile(nameFile, memoryStream.ToArray(), pathDirectory)
                End Using
            End Using
        End Sub

        ''' <summary>
        ''' Сжимает данные методом Deflate.
        ''' </summary>
        ''' <param name="data">Массив данных</param>
        Public Shared Function Compress(data As Byte()) As Byte()
            Dim output = New MemoryStream()
            Using deflate = New Compression.DeflateStream(output, Compression.CompressionMode.Compress, True)
                deflate.Write(data, 0, data.Length)
                deflate.Close()
            End Using
            Return output.ToArray()
        End Function

        ''' <summary>
        ''' Выполняет архивацию в указанное место.
        ''' </summary>
        ''' <param name="pathArchive">Путь к архиву</param>
        Public Sub Archive(pathArchive As String)
            Dim result = New List(Of Byte)
            Dim endRecord = New ZipEndRecord

            For Each f In _filesZip
                f.CentralDirectory.OffsetHeader = result.Count
                result.AddRange(f.ToByteListFile)
            Next
            endRecord.StartCDOffset = result.Count
            For Each f In _filesZip
                result.AddRange(f.ToByteListDirectory)
                endRecord.SizeCentralDir += f.ToByteListDirectory.Count
            Next

            endRecord.TotalEntriesDisk = _filesZip.Count
            endRecord.TotalEntries = _filesZip.Count
            result.AddRange(endRecord.ToByteList)

            File.WriteAllBytes(pathArchive, result.ToArray)
        End Sub

        ''' <summary>
        ''' Выполняет архивацию возвращая массив байтов.
        ''' </summary>
        Public Function Archive() As Byte()
            Dim result = New List(Of Byte)
            Dim endRecord = New ZipEndRecord

            For Each f In _filesZip
                f.CentralDirectory.OffsetHeader = result.Count
                result.AddRange(f.ToByteListFile)
            Next
            endRecord.StartCDOffset = result.Count
            For Each f In _filesZip
                result.AddRange(f.ToByteListDirectory)
                endRecord.SizeCentralDir += f.ToByteListDirectory.Count
            Next

            endRecord.TotalEntriesDisk = _filesZip.Count
            endRecord.TotalEntries = _filesZip.Count
            result.AddRange(endRecord.ToByteList)

            Return result.ToArray
        End Function

    End Class

End Namespace

