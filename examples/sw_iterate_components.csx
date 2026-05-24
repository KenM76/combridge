// sw_iterate_components.csx
//
// List every top-level component in the active SolidWorks assembly with its
// file path, configuration, and suppression/visibility state. Demonstrates
// IAssemblyDoc.GetComponents.
//
// API surface (verified against the canonical SW RAG):
//   - IAssemblyDoc.GetComponents(bool ToplevelOnly) → object[]
//   - IComponent2.Name2, .ReferencedConfiguration, .GetPathName()
//   - IComponent2.IsSuppressed(), .IsHidden(bool ConsiderSuppressed)
//   - IComponent2.GetSuppression2() → swComponentSuppressionState_e
//     (NOT GetSuppression — that's marked obsolete in SW 2024+)
// Enum values (swComponentSuppressionState_e):
//   0=Suppressed, 1=Lightweight, 2=FullyResolved, 3=Resolved,
//   4=FullyLightweight, 5=InternalIdMismatch
//
// Run:  combridge solidworks run-script examples\sw_iterate_components.csx -

if (swAssy is null)
{
    Console.Error.WriteLine($"Active document is not an assembly (type={swDocType}). Open an .SLDASM first.");
    return 2;
}

// false = top-level only (pass true to recursively descend into sub-assemblies).
object[] comps = (object[])swAssy.GetComponents(false);

Console.WriteLine($"Assembly: {swDoc?.GetTitle()}");
Console.WriteLine($"Top-level components: {comps.Length}");
Console.WriteLine();

int suppressed = 0, lightweight = 0, resolved = 0, other = 0;
for (int i = 0; i < comps.Length; i++)
{
    var c = (IComponent2)comps[i];

    bool isSup = c.IsSuppressed();
    bool isHid = !isSup && c.IsHidden(false);
    string flag = isSup ? "[S]" : isHid ? "[H]" : "[ ]";
    string cfg  = string.IsNullOrEmpty(c.ReferencedConfiguration) ? "" : $"@{c.ReferencedConfiguration}";

    Console.WriteLine($"  {flag} {c.Name2,-50}  {cfg}");
    Console.WriteLine($"       → {c.GetPathName()}");

    // Bucket by resolved/lightweight/suppressed for the summary.
    switch (c.GetSuppression2())
    {
        case 0:                                  suppressed++; break;
        case 1: case 4:                          lightweight++; break;
        case 2: case 3:                          resolved++; break;
        default:                                 other++; break;
    }
}

Console.WriteLine();
Console.WriteLine($"Summary: {comps.Length} total — "
                + $"{resolved} resolved, {lightweight} lightweight, {suppressed} suppressed"
                + (other > 0 ? $", {other} other state" : ""));
