using System;

namespace Errepar.MetadataManager.Process.Models
{
    public class WeeklyWindow
    {
        public DayOfWeek Day { get; set; }
        public TimeOnly Start { get; set; }
        public TimeOnly End { get; set; }

        public WeeklyWindow(DayOfWeek day, TimeOnly start, TimeOnly end)
        {
            Day = day;
            Start = start;
            End = end;
        }
    }
}
