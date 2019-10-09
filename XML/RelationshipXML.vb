Imports OpenXML.Relationship

Namespace XML

    ''' <summary>
    ''' XML-элемент связи в формате OfficeOpenXML.
    ''' </summary>
    Class RelationshipXML
        Inherits XElement

        ''' <summary>
        ''' Создаёт XML-элемент связи в файле с расширением .rels для формата OpenXML.
        ''' </summary>
        ''' <param name="rel">Связь</param>
        Sub New(rel As IRelationship)
            MyBase.New(NS.rs + "Relationship",
                       New XAttribute("Id", rel.Id),
                       New XAttribute("Type", rel.Type),
                       New XAttribute("Target", rel.Target))
            If rel.TargetMode IsNot Nothing Then
                Add(New XAttribute("TargetMode", rel.TargetMode))
            End If
        End Sub

    End Class

End Namespace

