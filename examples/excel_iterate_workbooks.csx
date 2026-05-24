// excel_iterate_workbooks.csx
//
// List every open Excel workbook and each workbook's worksheets, with the
// used-range address for each sheet. Useful as a quick "what's loaded?"
// sanity check.
//
// Run:  combridge excel run-script examples\excel_iterate_workbooks.csx -

Console.WriteLine($"Excel {xlApp.Version}   (Workbooks: {xlApp.Workbooks.Count})");
Console.WriteLine();

for (int wi = 1; wi <= xlApp.Workbooks.Count; wi++)
{
    Workbook wb = xlApp.Workbooks[wi];
    string saved = string.IsNullOrEmpty(wb.Path) ? "(unsaved)" : wb.FullName;
    Console.WriteLine($"[{wi}] {wb.Name}   path={saved}   sheets={wb.Worksheets.Count}");

    for (int si = 1; si <= wb.Worksheets.Count; si++)
    {
        Worksheet sh = wb.Worksheets[si];
        string addr;
        try { addr = sh.UsedRange.Address[false, false]; } catch { addr = "(empty)"; }
        Console.WriteLine($"     ({si}) {sh.Name,-30}  used={addr}");
    }
}
