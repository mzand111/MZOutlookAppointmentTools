using System;
using System.Collections.Generic;
using System.Linq;

namespace MZOutlookAppointmentTools.iCalendarTools
{
    public partial class RecurrenceStringTools
    {

        public static bool AreEqual(string recurrencePattern1, string recurrencePattern2)
        {
            if (string.IsNullOrWhiteSpace(recurrencePattern1))
            {
                if (string.IsNullOrWhiteSpace(recurrencePattern2))
                {
                    return true;
                }
                return false;
            }
            if (string.IsNullOrWhiteSpace(recurrencePattern2))
            {
                if (string.IsNullOrWhiteSpace(recurrencePattern1))
                {
                    return true;
                }
                return false;
            }
            var baseItemParts = recurrencePattern1.Split(';');
            var targetItemParts = recurrencePattern2.Split(';');
            Dictionary<string, string> baseKeyValuePairs = new Dictionary<string, string>();
            foreach (var baseItemPart in baseItemParts)
            {
                if (string.IsNullOrWhiteSpace(baseItemPart))
                    continue;
                var tmp = baseItemPart.Split('=');

                var basePartKey = tmp[0].Trim().ToUpperInvariant();
                var basePartValue = tmp[1].Trim().ToUpperInvariant();
                if (basePartKey == "BYDAY")
                {
                    baseKeyValuePairs.Add(basePartKey, SortDaysOfWeek(basePartValue));
                }
                else
                {
                    baseKeyValuePairs.Add(basePartKey, basePartValue);
                }
            }
            foreach (var targetItemPart in targetItemParts)
            {
                if (string.IsNullOrWhiteSpace(targetItemPart))
                    continue;
                var tmp = targetItemPart.Split('=');
                if (tmp.Length != 2)
                {
                    throw new Exception($"Error in parsing this part: {targetItemPart}");
                }
                var targetPartKey = tmp[0].Trim().ToUpperInvariant();
                if (!baseKeyValuePairs.ContainsKey(targetPartKey))
                {
                    return false;
                }
                var targetPartValue = tmp[1].Trim().ToUpperInvariant();
                if (targetPartKey == "BYDAY")
                {
                    var sortedByDay = SortDaysOfWeek(targetPartValue);
                    if (sortedByDay != baseKeyValuePairs[targetPartKey])
                    {
                        return false;
                    }
                }
                else
                {
                    if (targetPartValue != baseKeyValuePairs[targetPartKey])
                    {
                        return false;
                    }
                }
            }
            return true;
        }
        private static readonly string[] WeekdaysOrder = { "SU", "MO", "TU", "WE", "TH", "FR", "SA" };
        private static string SortDaysOfWeek(string input)
        {
            var days = input.Split(new[] { ',', ' ' }, StringSplitOptions.RemoveEmptyEntries)
                .Select(day => day.Trim()).OrderBy(day => Array.IndexOf(WeekdaysOrder, day)).ToArray(); return string.Join(",", days);
        }
    }
}
