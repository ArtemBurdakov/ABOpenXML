Namespace Zip

    ''' <summary>
    ''' Структура ZIP-архива, содержащая метаданные файла.
    ''' </summary>
    Class ZipLocalFileHeader
        Property Header As UInteger = &H04034B50
        Property Version As UShort = 20
        Property BitFlag As UShort = &H00
        Property CompMethod As UShort = &H00
        Property LastModTime As UShort = 0
        Property LastModDate As UShort = 0
        Property Crc32 As UInteger = 0
        Property CompSize As UInteger = 0
        Property UncompSize As UInteger = 0
        Property FileNameLength As UShort = 0
        Property ExtraFieldLength As UShort = 0

        ''' <summary>
        ''' Преобразует структуру в байтовый список.
        ''' </summary>
        Function ToByteList() As List(Of Byte)
            Dim result = New List(Of Byte)
            ConvertExtra.ToByteList(Header, result)
            ConvertExtra.ToByteList(Version, result)
            ConvertExtra.ToByteList(BitFlag, result)
            ConvertExtra.ToByteList(CompMethod, result)
            ConvertExtra.ToByteList(LastModTime, result)
            ConvertExtra.ToByteList(LastModDate, result)
            ConvertExtra.ToByteList(Crc32, result)
            ConvertExtra.ToByteList(CompSize, result)
            ConvertExtra.ToByteList(UncompSize, result)
            ConvertExtra.ToByteList(FileNameLength, result)
            ConvertExtra.ToByteList(ExtraFieldLength, result)
            Return result
        End Function

    End Class

End Namespace

