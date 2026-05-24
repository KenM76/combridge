# Scripting Cookbook — Per-Plugin Recipes

Common `.csx` patterns for each shipped plugin. Each recipe is a complete,
copy-pasteable script body (drop into a `.csx` file, run with
`combridge <plugin> run-script <file.csx> <out>`).

> Roslyn-script rules to remember:
> - Top-level statements only (no `class Program { static void Main }` wrapper)
> - `using var x = ...` declaration form NOT supported — use `using (var x = ...) { ... }` block form
> - Top-level `var` declarations hoist but assignments don't — order matters
> - `dynamic` works (Microsoft.CSharp is in the default refs)
> - `Console.Out` and `Console.Error` are redirected to the bridge's output writer

## Excel (`xlApp`, `xlBook`, `xlSheet`)

### Read a single cell

```csharp
var val = xlSheet?.Cells[1, 1]?.Value;
Console.WriteLine($"A1 = {val ?? "(empty)"}");
```

Range indexes are 1-based. `.Value` returns `object` — may be `double`,
`string`, `DateTime`, `null`, or `int`.

### Write a cell

```csharp
xlSheet.Cells[1, 1] = "Hello";
xlSheet.Cells[1, 2] = 42;
xlBook.Save();   // .SaveAs("path") for new files
```

### Iterate the used range

```csharp
var used = xlSheet.UsedRange;
int rows = used.Rows.Count;
int cols = used.Columns.Count;
Console.WriteLine($"Range: {rows} rows × {cols} cols");

// Bulk read — far faster than cell-by-cell for large ranges
var arr = used.Value as object[,];
if (arr is not null)
{
    for (int r = 1; r <= rows; r++)
    for (int c = 1; c <= cols; c++)
        Console.Write($"{arr[r, c]}\t");
    Console.WriteLine();
}
```

Single-cell ranges return scalar `Value`, multi-cell return `object[,]` —
always check the array cast.

### Iterate every open workbook + sheet

```csharp
for (int wi = 1; wi <= xlApp.Workbooks.Count; wi++)
{
    Workbook wb = xlApp.Workbooks[wi];
    Console.WriteLine($"\n[{wi}] {wb.Name}  ({wb.FullName})");
    for (int si = 1; si <= wb.Worksheets.Count; si++)
    {
        Worksheet sh = wb.Worksheets[si];
        Console.WriteLine($"   ({si}) {sh.Name}  used={sh.UsedRange.Address}");
    }
}
```

### Add a new workbook + save

```csharp
var wb = xlApp.Workbooks.Add();
var sh = (Worksheet)wb.Worksheets[1];
sh.Name = "Generated";
sh.Cells[1, 1] = "timestamp";
sh.Cells[1, 2] = DateTime.Now.ToString("yyyy-MM-dd HH:mm:ss");
wb.SaveAs(@"D:\out\new.xlsx", 51 /* xlOpenXMLWorkbook */);
```

### Find a cell by value

```csharp
var used = xlSheet.UsedRange;
var hit = used.Find("target string", Type.Missing, XlFindLookIn.xlValues, XlLookAt.xlWhole);
if (hit is not null)
    Console.WriteLine($"Found at {hit.Address}");
```

### Filtering / autofilter

```csharp
xlSheet.UsedRange.AutoFilter(Field: 2, Criteria1: "Approved");
// later, restore: xlSheet.AutoFilterMode = false;
```

## Word (`wdApp`, `wdDoc`)

### Read whole document text

```csharp
if (wdDoc is null) { Console.WriteLine("No document open."); return 0; }
Console.WriteLine($"Length: {wdDoc.Content.Text.Length} chars");
Console.WriteLine(wdDoc.Content.Text);
```

### Find and replace

```csharp
var find = wdApp.Selection.Find;
find.Text = "OLD_TEXT";
find.Replacement.Text = "NEW_TEXT";
find.Execute(Replace: WdReplace.wdReplaceAll);
```

For document-scoped (not selection-scoped) replace, use `wdDoc.Content.Find`
in the same shape.

### Iterate paragraphs

```csharp
int n = 0;
foreach (Paragraph p in wdDoc.Paragraphs)
{
    n++;
    var text = p.Range.Text.TrimEnd('\r', '\n');
    Console.WriteLine($"[{n}] {text}");
    if (n >= 20) { Console.WriteLine("  ... (truncated)"); break; }
}
```

### Insert text at cursor

```csharp
wdApp.Selection.TypeText("Inserted at " + DateTime.Now);
wdApp.Selection.TypeParagraph();
```

### Save as PDF

```csharp
wdDoc.ExportAsFixedFormat(
    OutputFileName: @"D:\out\report.pdf",
    ExportFormat: WdExportFormat.wdExportFormatPDF);
```

## PowerPoint (`pptApp`, `pptPres`, `pptSlide`)

### Iterate slides + titles

```csharp
if (pptPres is null) { Console.WriteLine("No presentation open."); return 0; }
foreach (Slide s in pptPres.Slides)
{
    string title = "(no title)";
    try
    {
        var titlePh = s.Shapes.Title;
        if (titlePh?.HasTextFrame == MsoTriState.msoTrue)
            title = titlePh.TextFrame.TextRange.Text;
    }
    catch { /* slide may have no title placeholder */ }
    Console.WriteLine($"[{s.SlideIndex}] {title}");
}
```

### Add a slide

```csharp
var layout = pptPres.SlideMaster.CustomLayouts[1];   // first custom layout
var newSlide = pptPres.Slides.AddSlide(pptPres.Slides.Count + 1, layout);
if (newSlide.Shapes.Count > 0)
{
    var ph = newSlide.Shapes[1];
    if (ph.HasTextFrame == MsoTriState.msoTrue)
        ph.TextFrame.TextRange.Text = "Inserted slide";
}
```

### Export as PDF

```csharp
pptPres.ExportAsFixedFormat(
    Path: @"D:\out\deck.pdf",
    FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF);
```

## Outlook (`olApp`, `olNs`, `olExplorer`)

### List unread Inbox items

```csharp
var inbox = olNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
var items = inbox.Items.Restrict("[Unread] = true");
int n = 0;
foreach (object item in items)
{
    if (item is MailItem mail)
    {
        n++;
        Console.WriteLine($"[{n}] {mail.SentOn:yyyy-MM-dd HH:mm}  {mail.SenderEmailAddress}  → {mail.Subject}");
        if (n >= 25) { Console.WriteLine("  ... (truncated)"); break; }
    }
}
Console.WriteLine($"({n} unread shown)");
```

### Search by subject

```csharp
var inbox = olNs.GetDefaultFolder(OlDefaultFolders.olFolderInbox);
// jet-query syntax for Restrict
var hits = inbox.Items.Restrict("\"urn:schemas:httpmail:subject\" LIKE '%invoice%'");
foreach (object item in hits)
{
    if (item is MailItem m)
        Console.WriteLine($"{m.ReceivedTime:yyyy-MM-dd}  {m.Subject}");
}
```

### Send a mail

```csharp
var mail = (MailItem)olApp.CreateItem(OlItemType.olMailItem);
mail.To = "recipient@example.com";
mail.Subject = "Status update";
mail.Body = "Plain text body.";
// mail.HTMLBody = "<p>Or HTML.</p>";
mail.Send();   // queues; check Outbox if Outlook is offline
```

### List today's calendar items

```csharp
var cal = olNs.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);
cal.Items.IncludeRecurrences = true;     // expand recurring events
cal.Items.Sort("[Start]");
string today = DateTime.Today.ToString("MM/dd/yyyy");
string tomorrow = DateTime.Today.AddDays(1).ToString("MM/dd/yyyy");
var todays = cal.Items.Restrict($"[Start] >= '{today}' AND [Start] < '{tomorrow}'");
foreach (object item in todays)
{
    if (item is AppointmentItem appt)
        Console.WriteLine($"{appt.Start:HH:mm}-{appt.End:HH:mm}  {appt.Subject}");
}
```

Outlook restrict filters have quirky syntax — `[Start]` and `[Unread]` are
property accessors; date literals are single-quoted.

### List all mailbox stores (accounts)

```csharp
foreach (Store store in olNs.Stores)
    Console.WriteLine($"{store.StoreID,-8}  {store.DisplayName}  ({store.FilePath ?? "(server store)"})");
```

## SolidWorks (`swApp`, `swDoc`, `swPart`, `swAssy`, `swDrawing`)

> **⚠️ Crash-prone operations** — read `C:\personal_rag\solidworks\index.md`
> § Crashes before using `SetSuppression2`, `UpdateCutList`, `ForceRebuild3`,
> `ExportToDWG2`, or anything that mutates state in a loop. Several SW
> APIs have non-obvious crash modes that aren't in the official docs.

### Identify the active document

```csharp
if (swDoc is null) { Console.WriteLine("No active doc"); return 0; }
Console.WriteLine($"Title:    {swDoc.GetTitle()}");
Console.WriteLine($"Path:     {swDoc.GetPathName()}");
Console.WriteLine($"Type:     {swDocType}");
Console.WriteLine($"Config:   {swDoc.ConfigurationManager.ActiveConfiguration.Name}");
```

### Iterate assembly components (top-level)

```csharp
if (swAssy is null) { Console.WriteLine("Active doc isn't an assembly"); return 0; }
object[] comps = (object[])swAssy.GetComponents(false /* TopLevelOnly */);
foreach (var o in comps)
{
    var c = (IComponent2)o;
    Console.WriteLine($"{c.Name2}  → {c.GetPathName()}");
}
```

### Walk every feature on a part

```csharp
IFeature feat = swDoc.FirstFeature() as IFeature;
while (feat is not null)
{
    Console.WriteLine($"{feat.GetTypeName2(),-30}  {feat.Name}");
    feat = feat.GetNextFeature() as IFeature;
}
```

⚠️ Never call `ForceRebuild3` mid-walk — it invalidates the Feature RCW
pointers (see `C:\personal_rag\solidworks\lesson_20260424_forcerebuild3_invalidates_com_pointers.md`).

### Export part to STEP

```csharp
int errors = 0, warnings = 0;
bool ok = swDoc.Extension.SaveAs(
    @"D:\out\part.step",
    (int)swSaveAsVersion_e.swSaveAsCurrentVersion,
    (int)swSaveAsOptions_e.swSaveAsOptions_Silent,
    null /*ExportData*/,
    ref errors,
    ref warnings);
Console.WriteLine($"saved={ok}  errors={errors} warnings={warnings}");
```

### Save a screenshot of the active view

```csharp
swDoc.ShowNamedView2("*Isometric", -1);
swDoc.ViewZoomtofit2();
swDoc.SaveBMP(@"D:\out\screenshot.bmp", 1600, 1200);
```

## Cross-plugin patterns

### Multi-app workflow (library mode, NOT scripting)

You can't run `combridge` across two apps from a single .csx — each
run-script invocation attaches to one plugin. For multi-app workflows
(e.g. read Excel data + draw it in SW), use **library mode** (see
`LLM/consuming.md`) where your tool references `ComBridge.Core` and
attaches to multiple ProgIDs in one process.

### Output capture

`Console.WriteLine` is redirected to the writer passed to combridge's
output-file argument:

```
combridge excel run-script myscript.csx out.txt     # output → out.txt
combridge excel run-script myscript.csx -           # output → stdout
```

The script CANNOT bypass this redirect (e.g. `Debugger.Log` won't appear).

### Exit codes from scripts

A script can short-circuit with an explicit return value:

```csharp
if (xlBook is null) { Console.Error.WriteLine("No workbook"); return 2; }
// ... do work ...
return 0;
```

Returned `int` becomes the script-host's exit code (visible as `$?` /
`%ERRORLEVEL%` after `combridge` returns). Conventional codes:
0=success, non-zero=failure. The bridge maps script exceptions to its
own exit code 4.

### Handling missing globals

Every plugin's "active doc" global may be null (no doc open). Always:

```csharp
if (swDoc is null) { Console.Error.WriteLine("Open a doc first"); return 2; }
// proceed
```

This is true for `xlBook`, `xlSheet`, `wdDoc`, `pptPres`, `pptSlide`,
`swDoc`, `swPart`, `swAssy`, `swDrawing`. The `app` global itself is
always non-null (you wouldn't have reached the script otherwise).

### COM cleanup

Don't manually `Marshal.ReleaseComObject(...)` on globals or anything
you didn't explicitly get back from a factory. The bridge owns the
top-level RCW; the script exits and the runtime cleans up. Calling
ReleaseComObject too aggressively breaks the host's subsequent commands.
