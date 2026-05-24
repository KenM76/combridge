// powerpoint_list_slides.csx
//
// List every slide in the active presentation with its title (if any) and
// shape count. Demonstrates Slides collection iteration + safe title access.
//
// Run:  combridge powerpoint run-script examples\powerpoint_list_slides.csx -

if (pptPres is null) { Console.Error.WriteLine("No active PowerPoint presentation."); return 2; }

Console.WriteLine($"Presentation: {pptPres.Name}");
Console.WriteLine($"Slides:       {pptPres.Slides.Count}");
Console.WriteLine();

foreach (Slide s in pptPres.Slides)
{
    string title = "(no title)";
    try
    {
        // Shapes.Title can throw if the layout has no title placeholder.
        var ph = s.Shapes.Title;
        if (ph?.HasTextFrame == MsoTriState.msoTrue)
        {
            title = ph.TextFrame.TextRange.Text.Replace("\r", " / ").Replace("\n", " / ");
            if (title.Length > 60) title = title.Substring(0, 57) + "...";
        }
    }
    catch { /* no title placeholder — fine */ }

    Console.WriteLine($"  [{s.SlideIndex,3}] shapes={s.Shapes.Count,3}  {title}");
}
