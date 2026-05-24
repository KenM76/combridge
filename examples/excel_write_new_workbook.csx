// excel_write_new_workbook.csx
//
// Create a brand-new workbook, write a small data table, and save it.
// Demonstrates Workbooks.Add, bulk array assignment to a Range, and SaveAs.
//
// Run:  combridge excel run-script examples\excel_write_new_workbook.csx -

// Defaults to %TEMP%\combridge_demo.xlsx. Edit if you want it elsewhere.
string outPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "combridge_demo.xlsx");

System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outPath)!);

var wb = xlApp.Workbooks.Add();
var sh = (Worksheet)wb.Worksheets[1];
sh.Name = "Demo";

// Header row
sh.Cells[1, 1] = "id";
sh.Cells[1, 2] = "name";
sh.Cells[1, 3] = "score";

// Bulk-write a 10x3 table in a single COM call (much faster than cell-by-cell)
var rows = new object[10, 3];
for (int i = 0; i < 10; i++)
{
    rows[i, 0] = i + 1;
    rows[i, 1] = $"item-{i + 1:D2}";
    rows[i, 2] = Math.Round(new Random(i).NextDouble() * 100, 2);
}
sh.Range["A2", "C11"].Value = rows;

// Auto-fit columns + bold header
sh.UsedRange.Columns.AutoFit();
sh.Range["A1", "C1"].Font.Bold = true;

// 51 = xlOpenXMLWorkbook (.xlsx)
wb.SaveAs(outPath, 51);
Console.WriteLine($"Saved → {outPath}");
Console.WriteLine($"Rows written: 10 + header.");
