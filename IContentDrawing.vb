Imports OpenXML.Zip

''' <summary>
''' Контент графики для формата OfficeOpenXML.
''' </summary>
Public Interface IContentDrawing

    ReadOnly Property Type As String
    Property Anchor As IAnchor
    Property IdContent As Integer
    Property Id As String
    Property Name As String
    Function AddZip(zip As ZipFile) As ZipFile

End Interface

