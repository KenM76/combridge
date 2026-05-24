// word_doc_stats.csx
//
// Print common statistics about the active Word document: word/character/
// paragraph/page counts, line count, and section count. Demonstrates
// WdStatistic enum + Document.ComputeStatistics.
//
// Run:  combridge word run-script examples\word_doc_stats.csx -

if (wdDoc is null) { Console.Error.WriteLine("No active Word document."); return 2; }

Console.WriteLine($"Document:  {wdDoc.Name}");
Console.WriteLine($"Path:      {(string.IsNullOrEmpty(wdDoc.Path) ? "(unsaved)" : wdDoc.FullName)}");
Console.WriteLine();
Console.WriteLine($"Paragraphs: {wdDoc.Paragraphs.Count,8:N0}");
Console.WriteLine($"Sections:   {wdDoc.Sections.Count,8:N0}");
Console.WriteLine($"Tables:     {wdDoc.Tables.Count,8:N0}");
Console.WriteLine();

// ComputeStatistics returns counts that need a layout pass; reliable on a
// document that's been displayed at least once.
int words   = wdDoc.ComputeStatistics(WdStatistic.wdStatisticWords);
int chars   = wdDoc.ComputeStatistics(WdStatistic.wdStatisticCharacters);
int lines   = wdDoc.ComputeStatistics(WdStatistic.wdStatisticLines);
int pages   = wdDoc.ComputeStatistics(WdStatistic.wdStatisticPages);
int parasS  = wdDoc.ComputeStatistics(WdStatistic.wdStatisticParagraphs);

Console.WriteLine($"Words:      {words,8:N0}");
Console.WriteLine($"Characters: {chars,8:N0}");
Console.WriteLine($"Lines:      {lines,8:N0}");
Console.WriteLine($"Pages:      {pages,8:N0}");
Console.WriteLine($"(Statistics-API paragraph count: {parasS,8:N0})");
