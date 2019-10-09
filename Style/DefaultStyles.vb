Namespace Style

    ''' <summary>
    ''' Основной стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleHeaderDefault
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "HeaderDefault"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center, True)
            Font = New FontStyle("Calibri", 11, "FF000000", True, False, False)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Основной стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowFirstDefault
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowFirstDefault"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.General, AlignmentVerticalStyle.Bottom, True)
            Font = New FontStyle("Calibri", 11)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Основной стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowFirstDateDefault
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowFirsDatetDefault"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.General, AlignmentVerticalStyle.Bottom, True)
            Font = New FontStyle("Calibri", 11)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
            NumFmt = New NumFmtStyle("dd.mm.yyyy")
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Основной стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowSecondDefault
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowSecondDefault"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.General, AlignmentVerticalStyle.Bottom, True)
            Font = New FontStyle("Calibri", 11)
            Fill = New FillStyle(PatternTypeStyle.Solid, "FFBBBBBB")
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Основной стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowSecondDateDefault
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowSecondDateDefault"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.General, AlignmentVerticalStyle.Bottom, True)
            Font = New FontStyle("Calibri", 11)
            Fill = New FillStyle(PatternTypeStyle.Solid, "FFBBBBBB")
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
            NumFmt = New NumFmtStyle("dd.mm.yyyy")
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Тестовый стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleHeaderTest
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "HeaderTest"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center, True)
            Font = New FontStyle("Arial", 12, "FF4499CC", True, False, False)
            Dim gf = New GradientFill(0)
            gf.AddStop("FF22BB99", 0)
            gf.AddStop("FF00FF00", 1)
            Fill = New FillStyle(gf)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thick
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thick
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thick
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thick
            Border.BottomColor = New Color With {.RGB = "FF000000"}
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Тестовый стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowFirstTest
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowFirstTest"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center, True)
            Font = New FontStyle("Arial", 11, "FF4499CC", True, False, False)
            Dim gf = New GradientFill(90)
            gf.AddStop("FF22BB99", 0)
            gf.AddStop("FFFF0000", 1)
            Fill = New FillStyle(gf)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Тестовый стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowFirstDateTest
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowFirstDateTest"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center, True)
            Font = New FontStyle("Arial", 11, "FF4499CC", True, False, False)
            Dim gf = New GradientFill(90)
            gf.AddStop("FF22BB99", 0)
            gf.AddStop("FFFF0000", 1)
            Fill = New FillStyle(gf)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
            NumFmt = New NumFmtStyle("dd.mm.yyyy")
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Тестовый стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowSeconTest
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowSecondTest"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center, True)
            Font = New FontStyle("Arial", 11, "FF4499CC", True, False, False)
            Dim gf = New GradientFill(45)
            gf.AddStop("FF22BB99", 0)
            gf.AddStop("FF0000FF", 0.5)
            gf.AddStop("FFFF0000", 1)
            Fill = New FillStyle(gf)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
            NumFmt = New NumFmtStyle("0.00")
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

    ''' <summary>
    ''' Тестовый стиль ячейки в Excel для формата OfficeOpenXML.
    ''' </summary>
    Public Class StyleRowSecondDateTest
        Implements IStyleCell

        Sub New()
            Id = 0
            Type = 1
            ParentId = 0
            Name = "RowSecondDateTest"
            Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center, True)
            Font = New FontStyle("Arial", 11, "FF4499CC", True, False, False)
            Dim gf = New GradientFill(45)
            gf.AddStop("FF22BB99", 0)
            gf.AddStop("FF0000FF", 0.5)
            gf.AddStop("FFFF0000", 1)
            Fill = New FillStyle(gf)
            Border = New BorderStyle
            Border.LeftStyle = BorderLineStyles.Thin
            Border.LeftColor = New Color With {.RGB = "FF000000"}
            Border.RightStyle = BorderLineStyles.Thin
            Border.RightColor = New Color With {.RGB = "FF000000"}
            Border.TopStyle = BorderLineStyles.Thin
            Border.TopColor = New Color With {.RGB = "FF000000"}
            Border.BottomStyle = BorderLineStyles.Thin
            Border.BottomColor = New Color With {.RGB = "FF000000"}
            NumFmt = New NumFmtStyle("dd.mm.yyyy")
        End Sub

        Public Property Id As Integer Implements IStyleCell.Id
        Public Property Type As Integer Implements IStyleCell.Type
        Public Property ParentId As Integer Implements IStyleCell.ParentId
        Public Property Name As String Implements IStyleCell.Name
        Public Property Alignment As AlignmentStyle Implements IStyleCell.Alignment
        Public Property Font As FontStyle Implements IStyleCell.Font
        Public Property Fill As FillStyle Implements IStyleCell.Fill
        Public Property Border As BorderStyle Implements IStyleCell.Border
        Public Property NumFmt As NumFmtStyle Implements IStyleCell.NumFmt

    End Class

End Namespace

