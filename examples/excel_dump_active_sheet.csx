// Example: dump the used range of the active Excel sheet.
// Run with:
//   combridge excel run-script examples\excel_dump_active_sheet.csx out.tsv
if (xlSheet is null)
{
    Console.WriteLine("No active worksheet (active sheet may be a Chart).");
    return;
}
Console.WriteLine($"Sheet: {xlSheet.Name}");
var used = xlSheet.UsedRange;
Console.WriteLine($"Used range: {used.Address[false, false]}  ({used.Rows.Count} x {used.Columns.Count})");

var raw = used.Value;
if (raw is object[,] arr)
{
    int rows = arr.GetLength(0);
    int cols = arr.GetLength(1);
    for (int r = 1; r <= rows; r++)
    {
        for (int c = 1; c <= cols; c++)
        {
            if (c > 1) Console.Write('\t');
            Console.Write(arr[r, c]?.ToString() ?? "");
        }
        Console.WriteLine();
    }
}
else
{
    Console.WriteLine(raw?.ToString() ?? "");
}
