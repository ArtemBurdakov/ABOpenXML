Namespace Zip

    ''' <summary>
    ''' Структура ZIP-архива, содержащая информацию о файлах и дерикториях.
    ''' </summary>
    Class ZipEndRecord
        Property Header As UInteger = &H06054B50
        Property DiskNumber As UShort = 0
        Property DiskNumberCD As UShort = 0
        Property TotalEntriesDisk As UShort = 0
        Property TotalEntries As UShort = 0
        Property SizeCentralDir As UInteger = 0
        Property StartCDOffset As UInteger = 0
        Property FileCommentLength As UShort = 0

        ''' <summary>
        ''' Преобразует структуру в байтовый список.
        ''' </summary>
        Function ToByteList() As List(Of Byte)
            Dim result = New List(Of Byte)
            ConvertExtra.ToByteList(Header, result)
            ConvertExtra.ToByteList(DiskNumber, result)
            ConvertExtra.ToByteList(DiskNumberCD, result)
            ConvertExtra.ToByteList(TotalEntriesDisk, result)
            ConvertExtra.ToByteList(TotalEntries, result)
            ConvertExtra.ToByteList(SizeCentralDir, result)
            ConvertExtra.ToByteList(StartCDOffset, result)
            ConvertExtra.ToByteList(FileCommentLength, result)
            Return result
        End Function

    End Class

End Namespace

