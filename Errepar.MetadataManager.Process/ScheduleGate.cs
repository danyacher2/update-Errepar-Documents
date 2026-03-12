using Errepar.MetadataManager.Process.Models;
using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;

public static class ScheduleGate
{
    public static (List<WeeklyWindow> allow, List<WeeklyWindow> deny, int rulesCount) LoadWindows(
        ClientContext context,
        string listTitle)
    {
        var list = context.Web.Lists.GetByTitle(listTitle);
        var query = new CamlQuery { ViewXml = "<View></View>" };
        var scheduleItems = list.GetItems(query);

        context.Load(scheduleItems);
        context.ExecuteQuery();

        var allow = new List<WeeklyWindow>();
        var deny = new List<WeeklyWindow>();

        foreach (var scheduleItem in scheduleItems)
        {
            if (scheduleItem["day"] is null)
                continue;

            var day = Enum.Parse<DayOfWeek>(scheduleItem["day"].ToString());
            var type = scheduleItem["allow"]?.ToString();

            var starts = ParseTimes(scheduleItem["startTime"]);
            var ends = ParseTimes(scheduleItem["endTime"]);

            for (int i = 0; i < Math.Min(starts.Count, ends.Count); i++)
            {
                var start = starts[i];
                var end = ends[i];

                if (end <= start)
                {
                    Console.WriteLine($"⚠ intervalo inválido {start} - {end}");
                    continue;
                }

                var window = new WeeklyWindow(day, start, end);
                if (type == "Allow")
                    allow.Add(window);
                else if (type == "Deny")
                    deny.Add(window);
            }
        }

        return (allow, deny, scheduleItems.Count);
    }

    public static Task WaitUntilAllowedAsync(
        IReadOnlyList<WeeklyWindow> allow,
        IReadOnlyList<WeeklyWindow> deny,
        TimeSpan checkEvery,
        CancellationToken cancellationToken = default)
    {
        return WaitCoreAsync(allow, deny, checkEvery, cancellationToken);
    }

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

    private static async Task WaitCoreAsync(
        IReadOnlyList<WeeklyWindow> allow,
        IReadOnlyList<WeeklyWindow> deny,
        TimeSpan checkEvery,
        CancellationToken cancellationToken)
    {
        while (!IsAllowed(DateTimeOffset.Now, allow, deny))
        {
            await Task.Delay(checkEvery, cancellationToken);
        }
    }

    private static List<TimeOnly> ParseTimes(object value)
    {
        if (value is string text)
            return new List<TimeOnly> { TimeOnly.Parse(text) };

        if (value is string[] stringArray)
            return stringArray.Select(TimeOnly.Parse).OrderBy(t => t).ToList();

        if (value is object[] objectArray)
            return objectArray.Select(v => TimeOnly.Parse(v?.ToString() ?? string.Empty)).OrderBy(t => t).ToList();

        return new List<TimeOnly>();
    }
}
