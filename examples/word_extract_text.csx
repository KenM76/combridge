// word_extract_text.csx
//
// Dump the full plain text of the active Word document to the output writer.
// Useful for piping document content to other tools or doing a quick
// "what's in this doc" check.
//
// Run:  combridge word run-script examples\word_extract_text.csx body.txt
// Or:   combridge word run-script examples\word_extract_text.csx -

if (wdDoc is null) { Console.Error.WriteLine("No active Word document."); return 2; }

// .Content.Text is the entire body excluding headers/footers/footnotes.
// Lines are separated by \r in Word's text; normalize to \n for downstream tools.
string text = wdDoc.Content.Text.Replace("\r", "\n");
Console.Write(text);
