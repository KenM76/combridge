// excel_find_value.csx
//
// Search the active sheet's used range for cells containing a literal value
// and print every hit. Demonstrates Range.Find / FindNext.
//
// Edit the `target` string below, then:
//   combridge excel run-script examples\excel_find_value.csx -

const string target = "TODO";   // ← edit me

if (xlSheet is null) { Console.Error.WriteLine("No active worksheet."); return 2; }

var used = xlSheet.UsedRange;
Range first = used.Find(target, Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlPart,
                        XlSearchOrder.xlByRows, XlSearchDirection.xlNext,
                        MatchCase: false);
if (first is null)
{
    Console.WriteLine($"No cells contain '{target}'.");
    return 0;
}

int hits = 0;
Range cur = first;
do
{
    hits++;
    Console.WriteLine($"  [{hits}] {cur.Address[false, false]}   value={cur.Value}");
    cur = used.FindNext(cur);
}
while (cur is not null && cur.Address != first.Address);

Console.WriteLine($"\nTotal: {hits} hit(s).");
