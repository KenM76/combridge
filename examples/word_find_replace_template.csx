// word_find_replace_template.csx
//
// Find-and-replace template. Edit the two strings below, then run against
// an open Word document. Demonstrates Selection.Find + Replace.
//
// SAFETY: this modifies the active document. Save first or work on a copy.
//
// Run:  combridge word run-script examples\word_find_replace_template.csx -

const string findText    = "OLD_STRING";   // ← edit me
const string replaceText = "NEW_STRING";   // ← edit me

if (wdDoc is null) { Console.Error.WriteLine("No active Word document."); return 2; }

// Use the document's Content range (not selection-scoped) for a doc-wide replace.
var find = wdDoc.Content.Find;
find.ClearFormatting();
find.Replacement.ClearFormatting();
find.Text = findText;
find.Replacement.Text = replaceText;

// Execute with Replace=All; returns true if anything matched.
bool any = find.Execute(
    Replace: WdReplace.wdReplaceAll,
    Forward: true,
    MatchCase: false,
    MatchWholeWord: false);

Console.WriteLine(any
    ? $"Replaced '{findText}' with '{replaceText}' throughout {wdDoc.Name}."
    : $"No occurrences of '{findText}' found.");
