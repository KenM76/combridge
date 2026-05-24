// sw_feature_walk.csx
//
// Walk every feature on the active SolidWorks document and print
// type + name + suppression state. Demonstrates FirstFeature / GetNextFeature.
//
// SAFETY: do NOT call ForceRebuild3, EditFeature, or any mutation API
// during this walk — they invalidate Feature RCW pointers mid-iteration.
// See C:\personal_rag\solidworks\lesson_20260424_forcerebuild3_invalidates_com_pointers.md.
//
// Run:  combridge solidworks run-script examples\sw_feature_walk.csx -

if (swDoc is null) { Console.Error.WriteLine("No active SolidWorks document."); return 2; }

Console.WriteLine($"Document: {swDoc.GetTitle()}  (type={swDocType})");
Console.WriteLine();

int total = 0, suppressed = 0;
IFeature? feat = swDoc.FirstFeature() as IFeature;

while (feat is not null)
{
    total++;
    string typeName = feat.GetTypeName2() ?? "(unknown)";
    string state    = feat.IsSuppressed() ? "[S]" : "[ ]";
    if (feat.IsSuppressed()) suppressed++;

    Console.WriteLine($"  {total,4}. {state} {typeName,-25} {feat.Name}");

    feat = feat.GetNextFeature() as IFeature;
}

Console.WriteLine();
Console.WriteLine($"Total features: {total}  (suppressed: {suppressed})");
