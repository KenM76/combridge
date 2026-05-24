// Example: print info about the active SolidWorks document.
// Run with:
//   combridge solidworks run-script examples\sw_active_doc.csx out.txt
if (swDoc is null)
{
    Console.WriteLine("No active document.");
    return;
}
Console.WriteLine($"Title: {swDoc.GetTitle()}");
Console.WriteLine($"Path:  {swDoc.GetPathName()}");
Console.WriteLine($"Type:  {swDocType}");
if (swPart is not null)
    Console.WriteLine("This is a Part document.");
else if (swAssy is not null)
    Console.WriteLine("This is an Assembly document.");
else if (swDrawing is not null)
    Console.WriteLine("This is a Drawing document.");
