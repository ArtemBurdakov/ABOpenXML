Imports OpenXML.Style

Namespace XML

    ''' <summary>
    ''' XML-элемент основного или прямого стиля в Excel для формата OfficeOpenXML.
    ''' </summary>
    Class XfXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент основного или прямого стиля в Excel для формата OpenXML.
        ''' </summary>
        ''' <param name="styleCell">Общая строка</param>
        Sub New(styleCell As IStyleCell)
            MyBase.New(NS.wb + "xf")

            If styleCell.Type = 1 Then
                Add(New XAttribute("xfId", styleCell.ParentId))
            End If

            If Not IsNothing(styleCell.Alignment) And styleCell.Type = 1 Then
                If styleCell.Alignment.Id > 0 Then
                    Add(New AlignmentStyleXML(styleCell.Alignment))
                    Add(New XAttribute("applyAlignment", 1))
                End If
            End If

            If Not IsNothing(styleCell.Font) Then
                Add(New XAttribute("fontId", styleCell.Font.Id))
                If styleCell.Font.Id > 0 Then Add(New XAttribute("applyFont", 1))
            End If

            If Not IsNothing(styleCell.Fill) Then
                Add(New XAttribute("fillId", styleCell.Fill.Id))
                If styleCell.Fill.Id > 0 Then Add(New XAttribute("applyFill", 1))
            End If

            If Not IsNothing(styleCell.Border) Then
                Add(New XAttribute("borderId", styleCell.Border.Id))
                If styleCell.Border.Id > 0 Then Add(New XAttribute("applyBorder", 1))
            End If

            If Not IsNothing(styleCell.NumFmt) Then
                Add(New XAttribute("numFmtId", styleCell.NumFmt.Id))
                If styleCell.NumFmt.Id > 0 Then Add(New XAttribute("applyNumberFormat", 1))
            Else
                Add(New XAttribute("numFmtId", 0))
            End If

        End Sub

    End Class

End Namespace

