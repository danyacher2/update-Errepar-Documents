using Errepar.MetadataManager.Process.Models;
using System;
using System.Collections.Generic;
using System.Linq;

public static class ScheduleGate
{
    public static bool IsAllowed(
        DateTimeOffset now,
        IReadOnlyList<WeeklyWindow> allow,
        IReadOnlyList<WeeklyWindow> deny)
    {
        var local = now.LocalDateTime;
        var day = local.DayOfWeek;
        var time = TimeOnly.FromDateTime(local);

        bool InWindow(WeeklyWindow w) =>
            w.Day == day && time >= w.Start && time < w.End;

        var denied = deny.Any(InWindow);
        if (denied) return false;

        if (allow.Count == 0) return true;
        return allow.Any(InWindow);
    }
}
