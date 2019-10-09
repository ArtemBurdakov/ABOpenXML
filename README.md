# ABOpenXML
OpenXML


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
