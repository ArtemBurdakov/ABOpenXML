Namespace Zip

    ''' <summary>
    ''' Структура ZIP-архива, содержащая расширенные метаданные файла.
    ''' </summary>
    Class ZipCentralDirectory
        Property Header As UInteger = &H02014B50
        Property MadeVer As UShort = &H3F
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
        Property FileCommentLength As UShort = 0
        Property DiskNumberStart As UShort = 0
        Property InternalFileAttr As UShort = 0
        Property ExternalFileAttr As UInteger = &H10
        Property OffsetHeader As UInteger = 0

        ''' <summary>
        ''' Преобразует структуру в байтовый список.
        ''' </summary>
        Function ToByteList() As List(Of Byte)
            Dim result = New List(Of Byte)
            ConvertExtra.ToByteList(Header, result)
            ConvertExtra.ToByteList(MadeVer, result)
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
            ConvertExtra.ToByteList(FileCommentLength, result)
            ConvertExtra.ToByteList(DiskNumberStart, result)
            ConvertExtra.ToByteList(InternalFileAttr, result)
            ConvertExtra.ToByteList(ExternalFileAttr, result)
            ConvertExtra.ToByteList(OffsetHeader, result)
            Return result
        End Function

    End Class

End Namespace

