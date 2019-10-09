Imports System.Text

Namespace Zip

    ''' <summary>
    ''' Файл в ZIP-архиве
    ''' </summary>
    Class FileZip
        Property LocalFileHeader As ZipLocalFileHeader
        Property Name As String
        Property Data As Byte()
        Property CentralDirectory As ZipCentralDirectory

        ''' <summary>
        ''' Создаёт файл в ZIP-архиве
        ''' </summary>
        ''' <param name="nameFile">Имя файла</param>
        ''' <param name="dataFile">Данные файла</param>
        Sub New(nameFile As String, Optional dataFile As Byte() = Nothing)
            Dim crcModel As RocksoftCrcModel
            LocalFileHeader = New ZipLocalFileHeader
            CentralDirectory = New ZipCentralDirectory
            Name = nameFile.Replace("\", "/")

            If Not IsNothing(dataFile) Then
                Data = ZipFile.Compress(dataFile)
                crcModel = New RocksoftCrcModel(32, &H04C11DB7)
                LocalFileHeader.Version = &H0A
                LocalFileHeader.CompMethod = &H08
                LocalFileHeader.Crc32 = crcModel.ComputeCrc(dataFile)
                LocalFileHeader.CompSize = Data.Length
                LocalFileHeader.UncompSize = dataFile.Length
                CentralDirectory.Version = &H0A
                CentralDirectory.CompMethod = &H08
                CentralDirectory.Crc32 = LocalFileHeader.Crc32
                CentralDirectory.CompSize = LocalFileHeader.CompSize
                CentralDirectory.UncompSize = LocalFileHeader.UncompSize
                CentralDirectory.ExternalFileAttr = &H20
            End If

            LocalFileHeader.LastModTime = Convert.ToUInt16(Date.Now.Hour) << 11 Or Convert.ToUInt16(Date.Now.Minute) << 5 Or Convert.ToUInt16(Date.Now.Second)
            LocalFileHeader.LastModDate = Convert.ToUInt16(Date.Now.Year - 1980) << 9 Or Convert.ToUInt16(Date.Now.Month) << 5 Or Convert.ToUInt16(Date.Now.Day)
            LocalFileHeader.FileNameLength = Name.Length
            CentralDirectory.LastModTime = LocalFileHeader.LastModTime
            CentralDirectory.LastModDate = LocalFileHeader.LastModDate
            CentralDirectory.FileNameLength = LocalFileHeader.FileNameLength
        End Sub

        ''' <summary>
        ''' Возвращает список байтов структуры LocalFileHeader, имени файла и данных.
        ''' </summary>
        Function ToByteListFile() As List(Of Byte)
            Dim result = New List(Of Byte)
            If Not IsNothing(LocalFileHeader) Then
                result.AddRange(LocalFileHeader.ToByteList)
            End If
            result.AddRange(Encoding.UTF8.GetBytes(Name))
            If Not IsNothing(Data) Then
                result.AddRange(Data)
            End If
            Return result
        End Function

        ''' <summary>
        ''' Возвращает список байтов структуры CentralDirectory и имени файла.
        ''' </summary>
        Function ToByteListDirectory() As List(Of Byte)
            Dim result = New List(Of Byte)
            If Not IsNothing(CentralDirectory) Then
                result.AddRange(CentralDirectory.ToByteList)
            End If
            result.AddRange(Encoding.UTF8.GetBytes(Name))
            Return result
        End Function

    End Class

End Namespace

