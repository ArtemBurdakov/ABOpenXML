# ABOpenXML

Example C#

```
using OpenXML;
using OpenXML.Style;

namespace Test
{
    class Program
    {

        static void Main(string[] args)
        {
            OfficeOpenXML ox = new OfficeOpenXML();
            WorkBook wb = ox.WorkBook;
            Sheet s1 = wb.AddSheet("Отчёт");

            IStyleCell styleHeader = new StyleCell();
            styleHeader.Alignment = new AlignmentStyle(AlignmentHorizontalStyle.Center, AlignmentVerticalStyle.Center);
            styleHeader.Font = new FontStyle(14, true, false, false);

            IStyleCell styleTableHeader = new StyleHeaderDefault();
            IStyleCell styleTable = new StyleRowFirstDefault();

            s1.SetColumnsWidth(1, 15);
            s1.SetColumnsWidth(2, 25);

            s1.SetMergeCell("A1:B2");

            s1.AddCell("A1", "Тест", styleHeader);

            s1.AddCell("A3", "Столбец 1", styleTableHeader);
            s1.AddCell("B3", "Столбец 2", styleTableHeader);

            for (int i = 4; i < 10; i++)
            {
                SheetRow row = new SheetRow { Number = i };
                row.AddCell(new SheetCell("A" + i, "dfklgj", styleTable));
                row.AddCell(new SheetCell("B" + i, 45, styleTable));
                s1.AddRow(row);
            }

            ox.Open(@"D:\Test.xlsx");
        }
    }
}

```

Example VB.net


```
Imports OpenXML
Imports OpenXML.Style

Module Module1

    Sub Main()
        Dim sw1 As New Stopwatch
        sw1.Start()
        CreateReport()
        sw1.Stop()
        Console.WriteLine(sw1.Elapsed)
        Console.ReadLine()
    End Sub

    Private Sub CreateReport()
        Dim ox = New OfficeOpenXML
        Dim wb = ox.WorkBook
        Dim s1 = wb.AddSheet("Отчёт")
        Dim indexRow = 3

        s1.SetPageMargins(0.3, 0.3, 0.75, 0.75, 0.25, 0.25)
        s1.SetPageSetup(PageOrientation.Portrait, 50)

        Dim headerStyle = New StyleCell
        headerStyle.Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center)
        headerStyle.Font = New FontStyle(14, True, False, False)

        Dim tableHeaderStyle = New StyleHeaderTest
        Dim tableStyle = New StyleRowFirstDefault
        Dim centerTableStyle = New StyleRowFirstDefault
        centerTableStyle.Alignment = New AlignmentStyle(AlignmentHorizontalStyle.Center)

        s1.SetColumnsWidth(1, 15)
        s1.SetColumnsWidth(2, 10)

        s1.SetMergeCell("A1:B1")

        s1.AddCell("A1", "Тест", headerStyle)
        s1.AddCell("A2", "Тест1", tableHeaderStyle)
        s1.AddCell("B2", "Тест2", tableHeaderStyle)

        Dim row As SheetRow
        For Each test1 In DataTest1.SetBigExample()
            row = New SheetRow With {.Number = indexRow}
            row.AddCell(New SheetCell("A" & indexRow, test1.Name, centerTableStyle))
            row.AddCell(New SheetCell("B" & indexRow, "", centerTableStyle))
            s1.AddRow(row)
            indexRow += 1

            For Each test2 In test1.Test2
                row = New SheetRow With {.Number = indexRow, .OutlineLevel = 1}
                row.AddCell(New SheetCell("A" & indexRow, "", tableStyle))
                row.AddCell(New SheetCell("B" & indexRow, test2.Number, tableStyle))
                s1.AddRow(row)
                indexRow += 1
            Next
        Next

        row = New SheetRow With {.Number = indexRow}
        row.AddCell(New SheetCell("A" & indexRow, "Итого:", tableStyle))
        row.AddCell(New SheetCell("B" & indexRow, 100, tableStyle))
        s1.AddRow(row)

        ox.Open("C:\Test.xlsx")
    End Sub

    Class DataTest1

        Shared Function SetExample() As DataTest1()
            Return {
                New DataTest1 With {.Name = "Test1", .Test2 = DataTest2.SetExample()},
                New DataTest1 With {.Name = "Test2", .Test2 = DataTest2.SetExample()},
                New DataTest1 With {.Name = "Test3", .Test2 = DataTest2.SetExample()},
                New DataTest1 With {.Name = "Test4", .Test2 = DataTest2.SetExample()},
                New DataTest1 With {.Name = "Test5", .Test2 = DataTest2.SetExample()}
            }
        End Function

        Shared Function SetBigExample() As DataTest1()
            Dim r = New Random(Date.Now.Millisecond)
            Dim result = New List(Of DataTest1)
            For i = 0 To 20000
                result.Add(New DataTest1 With {.Name = "Test" & r.Next(2000), .Test2 = DataTest2.SetExample()})
            Next
            Return result.ToArray
        End Function

        Property Name As String

        Property Test2 As DataTest2()

    End Class

    Class DataTest2

        Shared Function SetExample() As DataTest2()
            Return {
                New DataTest2 With {.Number = 1},
                New DataTest2 With {.Number = 2},
                New DataTest2 With {.Number = 3},
                New DataTest2 With {.Number = 4},
                New DataTest2 With {.Number = 5}
            }
        End Function

        Property Number As Integer

    End Class
    
End Module
```


