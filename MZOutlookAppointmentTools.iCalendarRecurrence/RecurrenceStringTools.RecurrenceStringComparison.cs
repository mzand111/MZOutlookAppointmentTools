﻿using System;
using System.Collections.Generic;
using System.Linq;

namespace MZOutlookAppointmentTools.iCalendarTools
{
    public partial class RecurrenceStringTools
    {
        /// <summary>
        /// Compares two recurrence pattern strings for equality.
        /// </summary>
        /// <param name="recurrencePattern1">The first recurrence pattern string.</param>
        /// <param name="recurrencePattern2">The second recurrence pattern string.</param>
        /// <returns>True if the recurrence patterns are equal, otherwise false.</returns>
        /// <exception cref="InvalidOperationException">Thrown when a pattern string cannot be parsed.</exception>
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
                if (tmp.Length != 2)
                    throw new InvalidOperationException($"Could not parse the key-pair value in the pattern string:'{baseItemPart}', The pattern was: '{recurrencePattern1}'");
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
            List<string> matchedItems = new List<string>();
            foreach (var targetItemPart in targetItemParts)
            {
                if (string.IsNullOrWhiteSpace(targetItemPart))
                    continue;
                var tmp = targetItemPart.Split('=');
                if (tmp.Length != 2)
                {
                    throw new InvalidOperationException($"Could not parse the key-pair value in the pattern string:'{targetItemPart}', The pattern was: '{recurrencePattern2}'");
                }
                var targetPartKey = tmp[0].Trim().ToUpperInvariant();
                var targetPartValue = tmp[1].Trim().ToUpperInvariant();
                if (!baseKeyValuePairs.ContainsKey(targetPartKey))
                {

                    if (targetPartKey == "INTERVAL" && double.Parse(targetPartValue) == 1)
                    {
                        matchedItems.Add(targetPartKey);
                        continue;
                    }
                    return false;
                }

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
                        //If the numeric values are the same, continue (e.g. 02 == 2)
                        if (double.TryParse(targetPartValue, out double tv))
                        {
                            if (double.TryParse(baseKeyValuePairs[targetPartKey], out double bv))
                            {
                                if (tv == bv)
                                {
                                    matchedItems.Add(targetPartKey);
                                    continue;
                                }
                            }
                        }
                        else if (targetPartKey == "BYSETPOS")
                        {
                            if (baseKeyValuePairs.ContainsKey("BYDAY") && (baseKeyValuePairs["FREQ"] == "MONTHLY" || baseKeyValuePairs["FREQ"] == "YEARLY"))
                            {
                                if (double.TryParse(targetPartValue, out double bv))
                                {
                                    if (1 == bv)
                                        continue;

                                }
                            }

                        }
                        else if (targetPartKey == "UNTIL" && baseKeyValuePairs["FREQ"] == "DAILY")
                        {
                            var sps = baseKeyValuePairs[targetPartKey].Split('T');
                            var spt = targetPartValue.Split('T');
                            if (sps.Length == 2 && spt.Length == 2)
                            {
                                if (sps[0] == spt[0] &&
                                    ((sps[1] == "000000Z" && spt[1] == "235959Z")
                                    ||
                                    (spt[1] == "000000Z" && sps[1] == "235959Z")
                                    ))
                                //This is caused by the fix we added
                                {
                                    matchedItems.Add(targetPartKey);
                                    continue;
                                }
                            }
                        }

                        return false;
                    }
                }
                matchedItems.Add(targetPartKey);
            }

            foreach (var kvp in baseKeyValuePairs.Keys)
            {
                if (!string.IsNullOrWhiteSpace(kvp))
                {
                    //WKST is not supported
                    if (kvp != "WKST" && !matchedItems.Contains(kvp))
                    {
                        if (kvp == "BYSETPOS")
                        {

                        }
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
