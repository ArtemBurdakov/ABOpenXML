Namespace Zip

    ''' <summary>
    ''' Дополнительные преобразования типов
    ''' </summary>
    Class ConvertExtra

        Shared Sub ToByteArray(value As UInteger, array As Byte(), start As Integer)
            array(start) = value And &H000000FF
            array(start + 1) = value >> 8 And &H000000FF
            array(start + 2) = value >> 16 And &H000000FF
            array(start + 3) = value >> 24 And &H000000FF
        End Sub

        Shared Sub ToByteArray(value As UShort, array As Byte(), start As Integer)
            array(start) = value And &H00FF
            array(start + 1) = value >> 8 And &H00FF
        End Sub

        Shared Sub ToByteList(value As UInteger, list As List(Of Byte))
            list.Add(value And &H000000FF)
            list.Add(value >> 8 And &H000000FF)
            list.Add(value >> 16 And &H000000FF)
            list.Add(value >> 24 And &H000000FF)
        End Sub

        Shared Sub ToByteList(value As UShort, list As List(Of Byte))
            list.Add(value And &H00FF)
            list.Add(value >> 8 And &H00FF)
        End Sub

    End Class

End Namespace

