// powerpoint_export_pdf.csx
//
// Export the active presentation to PDF. Demonstrates ExportAsFixedFormat.
//
// Edit the output path below, then:
//   combridge powerpoint run-script examples\powerpoint_export_pdf.csx -

// Defaults to %TEMP%\combridge_deck.pdf. Edit if you want it elsewhere.
string outPath = System.IO.Path.Combine(System.IO.Path.GetTempPath(), "combridge_deck.pdf");

if (pptPres is null) { Console.Error.WriteLine("No active PowerPoint presentation."); return 2; }

System.IO.Directory.CreateDirectory(System.IO.Path.GetDirectoryName(outPath)!);

pptPres.ExportAsFixedFormat(
    Path: outPath,
    FixedFormatType: PpFixedFormatType.ppFixedFormatTypePDF,
    Intent: PpFixedFormatIntent.ppFixedFormatIntentPrint,
    FrameSlides: MsoTriState.msoTriStateMixed,
    HandoutOrder: PpPrintHandoutOrder.ppPrintHandoutVerticalFirst,
    OutputType: PpPrintOutputType.ppPrintOutputSlides);

Console.WriteLine($"Exported {pptPres.Slides.Count} slides → {outPath}");
