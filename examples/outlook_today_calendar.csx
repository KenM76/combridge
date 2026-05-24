// outlook_today_calendar.csx
//
// List today's calendar appointments from the default Calendar folder.
// Demonstrates AppointmentItem access + Items.Restrict with date filtering.
//
// Run:  combridge outlook run-script examples\outlook_today_calendar.csx -

var cal = olNs.GetDefaultFolder(OlDefaultFolders.olFolderCalendar);

// CRITICAL: must enable recurrence expansion + sort by Start BEFORE Restrict,
// otherwise recurring meetings won't appear.
cal.Items.IncludeRecurrences = true;
cal.Items.Sort("[Start]");

string today    = DateTime.Today.ToString("MM/dd/yyyy HH:mm");
string tomorrow = DateTime.Today.AddDays(1).ToString("MM/dd/yyyy HH:mm");
string filter   = $"[Start] >= '{today}' AND [Start] < '{tomorrow}'";

Console.WriteLine($"Calendar items for {DateTime.Today:yyyy-MM-dd}:");
Console.WriteLine();

var todays = cal.Items.Restrict(filter);
int n = 0;
foreach (object item in todays)
{
    if (item is AppointmentItem appt)
    {
        n++;
        Console.WriteLine($"  {appt.Start:HH:mm}-{appt.End:HH:mm}   {appt.Subject}");
        if (!string.IsNullOrEmpty(appt.Location))
            Console.WriteLine($"                     @ {appt.Location}");
    }
}

if (n == 0) Console.WriteLine("  (nothing scheduled today)");
else Console.WriteLine($"\n{n} item(s).");
