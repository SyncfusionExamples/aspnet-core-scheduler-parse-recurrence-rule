using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;
using System.Globalization;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;


namespace ScheduleSample
{
    public class RecurrenceHelper
    {
        #region Internal Fields

        //internal static Dictionary<string, List<string>> ICalTimeZones = AddTimeZones();
        internal static List<string> valueList = new List<string>();
        //internal static Dictionary<string, RecurrenceProperties> RecPropertiesDict = new Dictionary<string, RecurrenceProperties>();
        internal static string COUNT;
        internal static string RECCOUNT;
        internal static string DAILY;
        internal static string WEEKLY;
        internal static string MONTHLY;
        internal static string YEARLY;
        internal static string INTERVAL;
        internal static string INTERVALCOUNT;
        internal static string BYSETPOS;
        internal static string BYSETPOSCOUNT;
        internal static string BYDAY;
        internal static string BYDAYVALUE;
        internal static string BYMONTHDAY;
        internal static string BYMONTHDAYCOUNT;
        internal static string BYMONTH;
        internal static string BYMONTHCOUNT;
        internal static int BYDAYPOSITION;
        internal static int BYMONTHDAYPOSITION;
        internal static int WEEKLYBYDAYPOS;
        internal static string WEEKLYBYDAY;
        internal static List<DateTime> exDateList = new List<DateTime>();
        internal static Nullable<DateTime> UNTIL;
        #endregion

        #region Methods

        public static IQueryable<DateTime> GetRecurrenceDateTimeCollection(string RRule, DateTime RecStartDate)
        {
            var filteredDates = GetRecurrenceDateCollection(RRule, RecStartDate, null, 43);
            return filteredDates;
        }
        public static IQueryable<DateTime> GetRecurrenceDateTimeCollection(string RRule, DateTime RecStartDate, int NeverCount)
        {
            var filteredDates = GetRecurrenceDateCollection(RRule, RecStartDate, null, NeverCount);
            return filteredDates;
        }

        public static IQueryable<DateTime> GetRecurrenceDateTimeCollection(string RRule, DateTime RecStartDate, string RecException)
        {
            var filteredDates = GetRecurrenceDateCollection(RRule, RecStartDate, RecException, 43);
            return filteredDates;
        }

        public static IQueryable<DateTime> GetRecurrenceDateTimeCollection(string RRule, DateTime RecStartDate, string RecException, int NeverCount)
        {
            var filteredDates = GetRecurrenceDateCollection(RRule, RecStartDate, RecException, NeverCount);
            return filteredDates;
        }

        public static string GetRuleSummary(string recurrenceRule)
        {
            if (recurrenceRule == null)
            {
                throw new ArgumentNullException(nameof(recurrenceRule));
            }

            string comma = ",", space = " ";
            string summary = "every";
            var seperator = new[] { ';' };
            string[] splitedRule = recurrenceRule.Split(seperator);
            var ruleSeperator = new[] { '=', ';', ',' };
            string[] ruleArray = recurrenceRule.Split(ruleSeperator);
            var currentCulture = CultureInfo.CurrentCulture; //Intl.GetCulture();
            string[] totalDays = currentCulture.DateTimeFormat.AbbreviatedDayNames;
            string[] totalMonths = currentCulture.DateTimeFormat.AbbreviatedMonthNames;
            int byDayPosition, byMonthDayPosition, weeklyByDayPos;
            string count, freq, recurrenceCount, daily, weekly, monthly, yearly, interval, intervalCount, until, untilValue, weeklyByDay, bySetPos, bySetPosCount, byDay, byDayValue, byMonthDay, byMonthDayCount, byMonth, byMonthCount;
            FindKeyIndex(
                ruleArray,
                out recurrenceCount,
                out freq,
                out daily,
                out weekly,
                out monthly,
                out yearly,
                out bySetPos,
                out bySetPosCount,
                out interval,
                out intervalCount,
                out until,
                out untilValue,
                out count,
                out byDay,
                out byDayValue,
                out byMonthDay,
                out byMonthDayCount,
                out byMonth,
                out byMonthCount,
                out weeklyByDay,
                out byMonthDayPosition,
                out byDayPosition);
            FindWeeklyRule(out splitedRule, out weeklyByDay, out weeklyByDayPos, splitedRule, currentCulture.DateTimeFormat.FirstDayOfWeek);
            if (!string.IsNullOrEmpty(intervalCount))
            {
                summary += space + intervalCount;
            }

            switch (freq)
            {
                case "DAILY":
                    summary += (space + "day(s)");
                    break;
                case "WEEKLY":
                    summary += space + ( "week(s)") + space + ( "on") + space;
                    string[] days = weeklyByDay.Split('=')[1].Split(',');
                    int index = 0;
                    foreach (string day in days)
                    {
                        int nthweekDay = GetWeekDay(day) - 1;
                        summary += totalDays[nthweekDay];
                        summary += ((days.Length - 1) == index) ? string.Empty : comma + space;
                        index++;
                    }

                    break;
                case "MONTHLY":
                    summary += space + ( "months(s)") + space + ( "on");
                    summary += GetMonthSummary(totalDays, byMonthDayCount, byDayValue, bySetPosCount);
                    break;
                case "YEARLY":
                    summary += space + ( "year(s)") + space + ("on") + space;
                    summary += totalMonths[int.Parse(byMonthCount, CultureInfo.InvariantCulture) - 1];
                    summary += GetMonthSummary(totalDays, byMonthDayCount, byDayValue, bySetPosCount);
                    break;
            }

            if (!string.IsNullOrEmpty(recurrenceCount))
            {
                summary += comma + space + recurrenceCount + space + ("time(s)");
            }
            else if (!string.IsNullOrEmpty(untilValue))
            {
                var tempDate = DateTime.ParseExact(untilValue, "yyyyMMddTHHmmssZ", CultureInfo.CurrentCulture);
                summary += comma + space + ( "until") + space + tempDate.Day + space + totalMonths[tempDate.Month - 1] + space + tempDate.Year;
            }

            return summary;
        }

        private static string GetMonthSummary(string[] days, string byMonthDayCount, string byDayValue, string bySetPosCount)
        {
            string[] weekPositions = new string[] { "First", "Second", "Third", "Fourth", "Last" };
            string summary = " ";
            if (!string.IsNullOrEmpty(byMonthDayCount))
            {
                summary += int.Parse(byMonthDayCount, CultureInfo.InvariantCulture);
            }
            else if (!string.IsNullOrEmpty(byDayValue))
            {
                int nthweekDay = GetWeekDay(byDayValue) -1;
                var pos = int.Parse(bySetPosCount, CultureInfo.InvariantCulture) - 1;
                var weekPos = pos > -1 ?  weekPositions[pos] : weekPositions[weekPositions.Length -1];
                summary += weekPos + " " + days[nthweekDay];
            }

            return summary;
        }


        public static IQueryable<DateTime> GetRecurrenceDateCollection(string RRule, DateTime RecStartDate, string RecException, int NeverCount)
        {
            List<DateTime> RecDateCollection = new List<DateTime>();
            DateTime startDate = RecStartDate;
            var ruleSeperator = new[] { '=', ';', ',' };
            var weeklySeperator = new[] { ';' };
            exDateList = new List<DateTime>();
            string[] ruleArray = RRule.Split(ruleSeperator);
            FindKeyIndex(ruleArray, RecStartDate);
            string[] weeklyRule = RRule.Split(weeklySeperator);
            FindWeeklyRule(weeklyRule);
            if (RecException != null)
            {
                FindExdateList(RecException);
            }
            if (exDateList.Contains(RecStartDate.Date))
            {
                return new List<DateTime> { RecStartDate.Date };
            }
            if (ruleArray.Length != 0 && RRule != "")
            {
                DateTime addDate = startDate;
                int recCount;
                int.TryParse(RECCOUNT, out recCount);

                DateTime? endDate = null;
                for (int i = 0; i < ruleArray.Length; i++)
                {
                    if (ruleArray[i].Contains("UNTIL"))
                    {
                        StringBuilder sb = new StringBuilder(ruleArray[i + 1]);
                        sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
                        DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
                                                               CultureInfo.InvariantCulture);
                        DateTime localDate = TimeZoneInfo.ConvertTimeFromUtc(value.DateTime, TimeZoneInfo.Local);
                        endDate = new DateTime(localDate.Year, localDate.Month, localDate.Day, 0, 0, 0);
                    }
                }
                #region DAILY
                if (DAILY == "DAILY")
                {
                
                    if ((ruleArray.Length > 4 && INTERVAL == "INTERVAL") || ruleArray.Length == 4)
                    {
                        int DyDayGap = ruleArray.Length == 4 ? 1 : int.Parse(INTERVALCOUNT);
                        if (recCount == 0 && UNTIL == null)
                        {
                            recCount = NeverCount;
                        }
                        if (recCount > 0)
                        {
                            for (int i = 0; i < recCount; i++)
                            {
                                RecDateCollection.Add(addDate.Date);
                                addDate = addDate.AddDays(DyDayGap);
                            }
                        }
                        else if (UNTIL != null)
                        {
                            bool IsUntilDateReached = false;
                            while (!IsUntilDateReached)
                            {
                                if (DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL)) <= 0)
                                {
                                    RecDateCollection.Add(addDate.Date);
                                    addDate = addDate.AddDays(DyDayGap);
                                }
                                int statusValue = DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL));
                                if (statusValue > 0)
                                {
                                    IsUntilDateReached = true;
                                }
                            }
                        }
                    }
                }
                #endregion

                #region WEEKLY
                else if (WEEKLY == "WEEKLY")
                {
                    int WyWeekGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;
                    bool isweeklyselected = weeklyRule[WEEKLYBYDAYPOS].Length > 6;
                    if (recCount == 0 && UNTIL == null)
                    {
                        recCount = NeverCount;
                    }
                    if (recCount > 0)
                    {
                        while (RecDateCollection.Count < recCount && isweeklyselected)
                        {
                            GetWeeklyDateCollection(addDate, weeklyRule, RecDateCollection);
                            addDate = addDate.DayOfWeek == DayOfWeek.Saturday ? addDate.AddDays(((WyWeekGap - 1) * 7) + 1) : addDate.AddDays(1);
                        }
                    }
                    else if (UNTIL != null)
                    {
                        bool IsUntilDateReached = false;
                        while (!IsUntilDateReached && isweeklyselected)
                        {
                            GetWeeklyDateCollection(addDate, weeklyRule, RecDateCollection);
                            addDate = addDate.DayOfWeek == DayOfWeek.Saturday ? addDate.AddDays(((WyWeekGap - 1) * 7) + 1) : addDate.AddDays(1);
                            int statusValue = DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL));
                            if (statusValue > 0)
                            {
                                IsUntilDateReached = true;
                            }
                        }
                    }
                }
                #endregion

                #region MONTHLY
                else if (MONTHLY == "MONTHLY")
                {
                    int MyMonthGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;
                    int position = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? 6 : BYMONTHDAYPOSITION;
                    if (BYMONTHDAY == "BYMONTHDAY")
                    {
                        int monthDate = int.Parse(BYMONTHDAYCOUNT);
                        if (monthDate <= 30)
                        {
                            int currDate = int.Parse(startDate.Day.ToString());
                            var temp = new DateTime(addDate.Year, addDate.Month, monthDate);
                            addDate = monthDate < currDate ? temp.AddMonths(1) : temp;
                            if (recCount == 0 && UNTIL == null)
                            {
                                recCount = NeverCount;
                            }
                            if (recCount > 0)
                            {
                                for (int i = 0; i < recCount; i++)
                                {
                                    addDate = GetByMonthDayDateCollection(addDate, RecDateCollection, monthDate, MyMonthGap);
                                }
                            }
                            else if (UNTIL != null)
                            {
                                bool IsUntilDateReached = false;
                                while (!IsUntilDateReached)
                                {
                                    addDate = GetByMonthDayDateCollection(addDate, RecDateCollection, monthDate, MyMonthGap);
                                    int statusValue = DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL));
                                    if (statusValue > 0)
                                    {
                                        IsUntilDateReached = true;
                                    }
                                }
                            }
                        }
                        else
                        {
                            if (recCount == 0 && UNTIL == null)
                            {
                                recCount = NeverCount;
                            }
                            if (recCount > 0)
                            {
                                for (int i = 0; i < recCount; i++)
                                {
                                    if (addDate.Day == startDate.Day)
                                    {
                                        RecDateCollection.Add(addDate.Date);
                                    }
                                    else
                                    {
                                        i = i - 1;
                                    }
                                    addDate = addDate.AddMonths(MyMonthGap);
                                    addDate = new DateTime(addDate.Year, addDate.Month, DateTime.DaysInMonth(addDate.Year, addDate.Month));
                                }
                            }
                            else if (UNTIL != null)
                            {
                                bool IsUntilDateReached = false;
                                while (!IsUntilDateReached)
                                {
                                    if (addDate.Day == startDate.Day)
                                    {
                                        RecDateCollection.Add(addDate.Date);
                                    }
                                    addDate = addDate.AddMonths(MyMonthGap);
                                    addDate = new DateTime(addDate.Year, addDate.Month, DateTime.DaysInMonth(addDate.Year, addDate.Month));
                                    int statusValue = DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL));
                                    if (statusValue > 0)
                                    {
                                        IsUntilDateReached = true;
                                    }
                                }
                            }
                        }
                    }

                    else if (BYDAY == "BYDAY")
                    {
                        if (recCount == 0 && UNTIL == null)
                        {
                            recCount = NeverCount;
                        }
                        if (recCount > 0)
                        {
                            while (RecDateCollection.Count < recCount)
                            {
                                var weekCount = MondaysInMonth(addDate);
                                var monthStart = new DateTime(addDate.Year, addDate.Month, 1);
                                DateTime weekStartDate = monthStart.AddDays(-(int)(monthStart.DayOfWeek));
                                var monthStartWeekday = (int)(monthStart.DayOfWeek);
                                int nthweekDay = GetWeekDay(BYDAYVALUE) - 1;
                                int nthWeek;
                                int bySetPos = 0;
                                int setPosCount;
                                int.TryParse(BYSETPOSCOUNT, out setPosCount);
                                if (monthStartWeekday <= nthweekDay)
                                {
                                    if (setPosCount < 1)
                                    {
                                        bySetPos = weekCount + setPosCount;
                                    }
                                    else
                                    {
                                        bySetPos = setPosCount;
                                    }
                                    if (setPosCount < 0)
                                    {
                                        nthWeek = bySetPos;
                                    }
                                    else
                                    {
                                        nthWeek = bySetPos - 1;
                                    }
                                }
                                else
                                {
                                    if (setPosCount < 0)
                                    {
                                        bySetPos = weekCount + setPosCount;
                                    }
                                    else
                                    {
                                        bySetPos = setPosCount;
                                    }
                                    nthWeek = bySetPos;
                                }
                                addDate = weekStartDate.AddDays((nthWeek) * 7);
                                addDate = addDate.AddDays(nthweekDay);
                                if (addDate.CompareTo(startDate.Date) < 0)
                                {
                                    addDate = addDate.AddMonths(1);
                                    continue;
                                }
                                if (weekCount == 6 && addDate.Day == 23)
                                {
                                    int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
                                    bool flag = true;
                                    if (addDate.Month == 2)
                                    {
                                        flag = false;
                                    }
                                    if (flag)
                                    {
                                        //addDate = addDate.AddDays(7);
                                        RecDateCollection.Add(addDate.Date);
                                    }
                                    addDate = addDate.AddMonths(MyMonthGap);
                                }
                                else if (weekCount == 6 && addDate.Day == 24)
                                {
                                    int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
                                    bool flag = true;
                                    if (addDate.AddDays(7).Day != days)
                                    {
                                        flag = false;
                                    }
                                    if (flag)
                                    {
                                        addDate = addDate.AddDays(7);
                                        RecDateCollection.Add(addDate.Date);
                                    }
                                    addDate = addDate.AddMonths(MyMonthGap);
                                }
                                else if (!(addDate.Day <= 23 && int.Parse(BYSETPOSCOUNT) == -1))
                                {
                                    RecDateCollection.Add(addDate.Date);
                                    addDate = addDate.AddMonths(MyMonthGap);
                                }
                            }
                        }
                        else if (UNTIL != null)
                        {
                            while (addDate <= endDate)
                            {
                                if (addDate > endDate)
                                {
                                    break;
                                }

                                var monthStart = new DateTime(addDate.Year, addDate.Month, 1, addDate.Hour, addDate.Minute, addDate.Second);
                                DateTime weekStartDate = monthStart.AddDays(-(int)monthStart.DayOfWeek);
                                var monthStartWeekday = (int)monthStart.DayOfWeek;
                                int nthweekDay = RecurrenceHelper.GetWeekDay(BYDAYVALUE) - 1;
                                int nthWeek;
                                if (monthStartWeekday <= nthweekDay)
                                {
                                    nthWeek = int.Parse(BYSETPOSCOUNT, CultureInfo.InvariantCulture) - 1;
                                }
                                else
                                {
                                    nthWeek = int.Parse(BYSETPOSCOUNT, CultureInfo.InvariantCulture);
                                }

                                if (int.Parse(BYSETPOSCOUNT, CultureInfo.InvariantCulture) == -1)
                                {
                                    var lastDate = ScheduleUtils.LastDateOfMonth(monthStart);
                                    addDate = ScheduleUtils.GetWeekFirstDate(lastDate, nthweekDay);
                                }
                                else
                                {
                                    addDate = weekStartDate.AddDays(nthWeek * 7);
                                    addDate = addDate.AddDays(nthweekDay);
                                    if (addDate >= endDate)
                                    {
                                        break;
                                    }
                                }

                                if (addDate.CompareTo(startDate.Date) < 0)
                                {
                                    addDate = addDate.AddMonths(1);
                                    continue;
                                }


                                RecDateCollection.Add(addDate);


                                if (int.Parse(BYSETPOSCOUNT, CultureInfo.InvariantCulture) == -1)
                                {
                                    addDate = new DateTime(addDate.Year, addDate.Month, 1, addDate.Hour, addDate.Minute, addDate.Second).AddMonths(MyMonthGap);
                                }
                                else
                                {
                                    addDate = addDate.AddMonths(MyMonthGap).AddDays(-(addDate.Day - 1));
                                }
                            }
                        }
                    }
                }
                #endregion

                #region YEARLY
                else if (YEARLY == "YEARLY")
                {
                    int YyYearGap = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? int.Parse(INTERVALCOUNT) : 1;
                    int position = ruleArray.Length > 4 && INTERVAL == "INTERVAL" ? 6 : BYMONTHDAYPOSITION;
                    if (BYMONTHDAY == "BYMONTHDAY")
                    {
                        int monthIndex = int.Parse(BYMONTHCOUNT);
                        int dayIndex = int.Parse(BYMONTHDAYCOUNT);
                        if (monthIndex > 0 && monthIndex <= 12)
                        {
                            int bound = DateTime.DaysInMonth(addDate.Year, monthIndex);
                            if (bound >= dayIndex)
                            {
                                var specificDate = new DateTime(addDate.Year, monthIndex, dayIndex);
                                if (specificDate.Date < addDate.Date)
                                {
                                    addDate = specificDate;
                                    addDate = addDate.AddYears(1);
                                }
                                else
                                {
                                    addDate = specificDate;
                                }
                                if (recCount == 0 && UNTIL == null)
                                {
                                    recCount = NeverCount;
                                }
                                if (recCount > 0)
                                {
                                    for (int i = 0; i < recCount; i++)
                                    {
                                        RecDateCollection.Add(addDate.Date);
                                        addDate = addDate.AddYears(YyYearGap);
                                    }
                                }
                                else if (UNTIL != null)
                                {
                                    bool IsUntilDateReached = false;
                                    while (!IsUntilDateReached)
                                    {
                                        RecDateCollection.Add(addDate.Date);
                                        addDate = addDate.AddYears(YyYearGap);
                                        int statusValue = DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL));
                                        if (statusValue > 0)
                                        {
                                            IsUntilDateReached = true;
                                        }
                                    }
                                }
                            }
                        }
                    }
                    else if (BYDAY == "BYDAY")
                    {
                        int monthIndex = int.Parse(BYMONTHCOUNT);
                        if (recCount == 0 && UNTIL == null)
                        {
                            recCount = NeverCount;
                        }
                        if (recCount > 0)
                        {
                            while (RecDateCollection.Count < recCount)
                            {
                                var weekCount = MondaysInMonth(addDate);
                                var monthStart = new DateTime(addDate.Year, monthIndex, 1);
                                DateTime weekStartDate = monthStart.AddDays(-(int)(monthStart.DayOfWeek));
                                var monthStartWeekday = (int)(monthStart.DayOfWeek);
                                int nthweekDay = GetWeekDay(BYDAYVALUE) - 1;
                                int nthWeek;
                                int bySetPos = 0;
                                int setPosCount;
                                int.TryParse(BYSETPOSCOUNT, out setPosCount);
                                if (monthStartWeekday <= nthweekDay)
                                {
                                    if (setPosCount < 1)
                                    {
                                        bySetPos = weekCount + setPosCount;
                                    }
                                    else
                                    {
                                        bySetPos = setPosCount;
                                    }
                                    if (setPosCount < 0)
                                    {
                                        nthWeek = bySetPos;
                                    }
                                    else
                                    {
                                        nthWeek = bySetPos - 1;
                                    }
                                }
                                else
                                {
                                    if (setPosCount < 0)
                                    {
                                        bySetPos = weekCount + setPosCount;
                                    }
                                    else
                                    {
                                        bySetPos = setPosCount;
                                    }
                                    nthWeek = bySetPos;
                                }
                                if (setPosCount == -1)
                                {
                                    var lastDate = ScheduleUtils.LastDateOfMonth(monthStart);
                                    addDate = ScheduleUtils.GetWeekFirstDate(lastDate, nthweekDay);
                                }
                                else
                                {
                                    addDate = weekStartDate.AddDays((nthWeek) * 7);
                                    addDate = addDate.AddDays(nthweekDay);
                                }
                                if (addDate.CompareTo(startDate.Date) < 0)
                                {
                                    addDate = addDate.AddYears(1);
                                    continue;
                                }
                                if (weekCount == 6 && addDate.Day == 23)
                                {
                                    int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
                                    bool flag = true;
                                    if (addDate.Month == 2)
                                    {
                                        flag = false;
                                    }
                                    if (flag)
                                    {
                                        addDate = addDate.AddDays(7);
                                        RecDateCollection.Add(addDate.Date);
                                    }
                                    addDate = addDate.AddYears(YyYearGap);
                                }
                                else if (weekCount == 6 && addDate.Day == 24)
                                {
                                    int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
                                    bool flag = true;
                                    if (addDate.AddDays(7).Day != days)
                                    {
                                        flag = false;
                                    }
                                    if (flag)
                                    {
                                        addDate = addDate.AddDays(7);
                                        RecDateCollection.Add(addDate.Date);
                                    }
                                    addDate = addDate.AddYears(YyYearGap);
                                }
                                else if (!(addDate.Day <= 23 && int.Parse(BYSETPOSCOUNT) == -1))
                                {
                                    RecDateCollection.Add(addDate.Date);
                                    addDate = addDate.AddYears(YyYearGap);
                                }
                            }
                        }
                        else if (UNTIL != null)
                        {
                            bool IsUntilDateReached = false;
                            while (!IsUntilDateReached)
                            {
                                var weekCount = MondaysInMonth(addDate);
                                var monthStart = new DateTime(addDate.Year, monthIndex, 1);
                                DateTime weekStartDate = monthStart.AddDays(-(int)(monthStart.DayOfWeek));
                                var monthStartWeekday = (int)(monthStart.DayOfWeek);
                                int nthweekDay = GetWeekDay(BYDAYVALUE) - 1;
                                int nthWeek;
                                int bySetPos = 0;
                                int setPosCount;
                                int.TryParse(BYSETPOSCOUNT, out setPosCount);
                                if (monthStartWeekday <= nthweekDay)
                                {
                                    if (setPosCount < 1)
                                    {
                                        bySetPos = weekCount + setPosCount;
                                    }
                                    else
                                    {
                                        bySetPos = setPosCount;
                                    }
                                    if (setPosCount < 0)
                                    {
                                        nthWeek = bySetPos;
                                    }
                                    else
                                    {
                                        nthWeek = bySetPos - 1;
                                    }
                                }
                                else
                                {
                                    if (setPosCount < 0)
                                    {
                                        bySetPos = weekCount + setPosCount;
                                    }
                                    else
                                    {
                                        bySetPos = setPosCount;
                                    }
                                    nthWeek = bySetPos;
                                }
                                if (setPosCount == -1)
                                {
                                    var lastDate = ScheduleUtils.LastDateOfMonth(monthStart);
                                    addDate = ScheduleUtils.GetWeekFirstDate(lastDate, nthweekDay);
                                }
                                else
                                {
                                    addDate = weekStartDate.AddDays((nthWeek) * 7);
                                    addDate = addDate.AddDays(nthweekDay);
                                }
                                if (addDate.CompareTo(startDate.Date) < 0)
                                {
                                    addDate = addDate.AddYears(1);
                                    continue;
                                }
                                if (DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL)) <= 0)
                                {
                                    if (weekCount == 6 && addDate.Day == 23)
                                    {
                                        int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
                                        bool flag = true;
                                        if (addDate.Month == 2)
                                        {
                                            flag = false;
                                        }
                                        if (flag)
                                        {
                                            addDate = addDate.AddDays(7);
                                            RecDateCollection.Add(addDate.Date);
                                        }
                                        addDate = addDate.AddYears(YyYearGap);
                                    }
                                    else if (weekCount == 6 && addDate.Day == 24)
                                    {
                                        int days = DateTime.DaysInMonth(addDate.Year, addDate.Month);
                                        bool flag = true;
                                        if (addDate.AddDays(7).Day != days)
                                        {
                                            flag = false;
                                        }
                                        if (flag)
                                        {
                                            addDate = addDate.AddDays(7);
                                            RecDateCollection.Add(addDate.Date);
                                        }
                                        addDate = addDate.AddYears(YyYearGap);
                                    }
                                    else if (!(addDate.Day <= 23 && int.Parse(BYSETPOSCOUNT) == -1))
                                    {
                                        RecDateCollection.Add(addDate.Date);
                                        addDate = addDate.AddYears(YyYearGap);
                                    }
                                }
                                int statusValue = DateTime.Compare(addDate.Date, Convert.ToDateTime(UNTIL));
                                if (statusValue > 0)
                                {
                                    IsUntilDateReached = true;
                                }
                            }
                        }
                    }
                }
                #endregion
            }
            var filteredDates = RecDateCollection.Except(exDateList).ToList();
            return filteredDates.AsQueryable();
        }

        public static int MondaysInMonth(DateTime thisMonth)
        {
            DateTime today = thisMonth;
            //extract the month
            int daysInMonth = DateTime.DaysInMonth(today.Year, today.Month);
            DateTime firstOfMonth = new DateTime(today.Year, today.Month, 1);
            //days of week starts by default as Sunday = 0
            int firstDayOfMonth = (int)firstOfMonth.DayOfWeek;
            int weeksInMonth = (int)Math.Ceiling((firstDayOfMonth + daysInMonth) / 7.0);
            return weeksInMonth;
        }

        private static void GetWeeklyDateCollection(DateTime addDate, string[] weeklyRule, List<DateTime> RecDateCollection)
        {
            switch (addDate.DayOfWeek)
            {
                case DayOfWeek.Sunday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("SU"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
                case DayOfWeek.Monday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("MO"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
                case DayOfWeek.Tuesday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("TU"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
                case DayOfWeek.Wednesday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("WE"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
                case DayOfWeek.Thursday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("TH"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
                case DayOfWeek.Friday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("FR"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
                case DayOfWeek.Saturday:
                    {
                        if (weeklyRule[WEEKLYBYDAYPOS].Contains("SA"))
                        {
                            RecDateCollection.Add(addDate.Date);
                        }
                        break;
                    }
            }
        }

        private static DateTime GetByMonthDayDateCollection(DateTime addDate, List<DateTime> RecDateCollection, int monthDate, int MyMonthGap)
        {
            if (addDate.Month == 2 && monthDate > 28)
            {
                addDate = new DateTime(addDate.Year, addDate.Month, DateTime.DaysInMonth(addDate.Year, 2));
                // RecDateCollection.Add(addDate.Date);
                addDate = addDate.AddMonths(MyMonthGap);
                addDate = new DateTime(addDate.Year, addDate.Month, monthDate);
            }
            else
            {
                RecDateCollection.Add(addDate.Date);
                addDate = addDate.AddMonths(MyMonthGap);
            }
            return addDate;
        }

        private static DateTime GetByDayDateValue(DateTime addDate, DateTime monthStart)
        {
            DateTime weekStartDate = monthStart.AddDays(-(int)(monthStart.DayOfWeek));
            var monthStartWeekday = (int)(monthStart.DayOfWeek);
            int nthweekDay = GetWeekDay(BYDAYVALUE) - 1;
            int nthWeek;
            if (monthStartWeekday <= nthweekDay)
            {
                nthWeek = int.Parse(BYSETPOSCOUNT) - 1;
            }
            else
            {
                nthWeek = int.Parse(BYSETPOSCOUNT);
            }
            addDate = weekStartDate.AddDays((nthWeek) * 7);
            addDate = addDate.AddDays(nthweekDay);
            return addDate;
        }

        public static List<DateTime> GetRecurrenceExceptionsDates(string ruleDate)
        {
            var exDates = ruleDate.Split(',');
            for (int i = 0; i < exDates.Length; i++)
            {
                StringBuilder sb = new StringBuilder(exDates[i]);
                sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
                DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
                                                       CultureInfo.InvariantCulture);
                DateTime localDate = TimeZoneInfo.ConvertTimeFromUtc(value.DateTime, TimeZoneInfo.Local);
                exDateList.Add(localDate);
            }
            return exDateList;
        }
        private static void FindExdateList(string ruleException)
        {
            exDateList = new List<DateTime>();
            var exDates = ruleException.Split(',');
            for (int i = 0; i < exDates.Length; i++)
            {
                StringBuilder sb = new StringBuilder(exDates[i]);
                sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
                DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
                                                       CultureInfo.InvariantCulture);
                DateTime localDate = TimeZoneInfo.ConvertTimeFromUtc(value.DateTime, TimeZoneInfo.Local);
                exDateList.Add(localDate);
            }
        }

        private static int GetWeekDay(string weekDay)
        {
            switch (weekDay)
            {
                case "SU":
                    {
                        return 1;
                    }
                case "MO":
                    {
                        return 2;
                    }
                case "TU":
                    {
                        return 3;
                    }
                case "WE":
                    {
                        return 4;
                    }
                case "TH":
                    {
                        return 5;
                    }
                case "FR":
                    {
                        return 6;
                    }
                case "SA":
                    {
                        return 7;
                    }
            }
            return 8;
        }

        private static void FindWeeklyRule(string[] weeklyRule)
        {
            for (int i = 0; i < weeklyRule.Length; i++)
            {
                if (weeklyRule[i].Contains("BYDAY"))
                {
                    WEEKLYBYDAY = weeklyRule[i];
                    WEEKLYBYDAYPOS = i;
                    break;
                }
            }
        }

        public static DateTime GetRecurrenceUntilDate(string RRule)
        {
            var ruleSeperator = new[] { '=', ';', ',' };
            var weeklySeperator = new[] { ';' };
            string[] ruleArray = RRule.Split(ruleSeperator);
            DateTime dtUntil = new DateTime();
            for (int i = 0; i < ruleArray.Length; i++)
            {
                if (ruleArray[i].Contains("COUNT"))
                {
                    COUNT = ruleArray[i];
                    RECCOUNT = ruleArray[i + 1];
                }

                if (ruleArray[i].Contains("UNTIL"))
                {
                    StringBuilder sb = new StringBuilder(ruleArray[i + 1]);
                    sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
                    DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
                                                           CultureInfo.InvariantCulture);
                    DateTime localDate = TimeZoneInfo.ConvertTimeFromUtc(value.DateTime, TimeZoneInfo.Local);
                    dtUntil = localDate;
                    ;
                }
            }

            return dtUntil;
        }

        private static void FindKeyIndex(string[] ruleArray, DateTime startDate)
        {
            RECCOUNT = "";
            DAILY = "";
            WEEKLY = "";
            MONTHLY = "";
            YEARLY = "";
            BYSETPOS = "";
            BYSETPOSCOUNT = "";
            INTERVAL = "";
            INTERVALCOUNT = "";
            COUNT = "";
            BYDAY = "";
            BYDAYVALUE = "";
            BYMONTHDAY = "";
            BYMONTHDAYCOUNT = "";
            BYMONTH = "";
            BYMONTHCOUNT = "";
            WEEKLYBYDAY = "";
            UNTIL = null;

            for (int i = 0; i < ruleArray.Length; i++)
            {
                if (ruleArray[i].Contains("COUNT"))
                {
                    COUNT = ruleArray[i];
                    RECCOUNT = ruleArray[i + 1];
                }

                if (ruleArray[i].Contains("UNTIL"))
                {
                    StringBuilder sb = new StringBuilder(ruleArray[i + 1]);
                    sb.Insert(4, '-'); sb.Insert(7, '-'); sb.Insert(13, ':'); sb.Insert(16, ':');
                    DateTimeOffset value = DateTimeOffset.ParseExact(sb.ToString(), "yyyy-MM-dd'T'HH:mm:ss'Z'",
                                                           CultureInfo.InvariantCulture);
                    DateTime localDate = TimeZoneInfo.ConvertTimeFromUtc(value.DateTime, TimeZoneInfo.Local);
                    UNTIL = new DateTime(localDate.Year, localDate.Month, localDate.Day, 0, 0, 0);
                }

                if (ruleArray[i].Contains("DAILY"))
                {
                    DAILY = ruleArray[i];
                }

                if (ruleArray[i].Contains("WEEKLY"))
                {
                    WEEKLY = ruleArray[i];
                }
                if (ruleArray[i].Contains("INTERVAL"))
                {
                    INTERVAL = ruleArray[i];
                    INTERVALCOUNT = ruleArray[i + 1];
                }
                if (ruleArray[i].Contains("MONTHLY"))
                {
                    MONTHLY = ruleArray[i];
                }
                if (ruleArray[i].Contains("YEARLY"))
                {
                    YEARLY = ruleArray[i];
                }
                if (ruleArray[i].Contains("BYSETPOS"))
                {
                    BYSETPOS = ruleArray[i];
                    var weekCount = MondaysInMonth(startDate);
                    BYSETPOSCOUNT = ruleArray[i + 1];
                }
                if (ruleArray[i].Contains("BYDAY"))
                {
                    BYDAYPOSITION = i;
                    BYDAY = ruleArray[i];
                    BYDAYVALUE = ruleArray[i + 1];
                }
                if (ruleArray[i].Contains("BYMONTHDAY"))
                {
                    BYMONTHDAYPOSITION = i;
                    BYMONTHDAY = ruleArray[i];
                    BYMONTHDAYCOUNT = ruleArray[i + 1];
                }
                if (ruleArray[i].Contains("BYMONTH"))
                {
                    BYMONTH = ruleArray[i];
                    BYMONTHCOUNT = ruleArray[i + 1];
                }
            }
        }


#pragma warning disable CA1021 // Avoid out parameters
        public static void FindKeyIndex(
                 string[] ruleArray,
                 out string recurrenceCount,
                 out string freq,
                 out string daily,
                 out string weekly,
                 out string monthly,
                 out string yearly,
                 out string bySetPos,
                 out string bySetPosCount,
                 out string interval,
                 out string intervalCount,
                 out string until,
                 out string untilValue,
                 out string count,
                 out string byDay,
                 out string byDayValue,
                 out string byMonthDay,
                 out string byMonthDayCount,
                 out string byMonth,
                 out string byMonthCount,
                 out string weeklyByDay,
                 out int byMonthDayPosition,
                 out int byDayPosition)
        {
            byMonthDayPosition = 0;
            byDayPosition = 0;
            recurrenceCount = string.Empty;
            freq = string.Empty;
            daily = string.Empty;
            weekly = string.Empty;
            monthly = string.Empty;
            yearly = string.Empty;
            bySetPos = string.Empty;
            bySetPosCount = string.Empty;
            interval = string.Empty;
            intervalCount = string.Empty;
            count = string.Empty;
            byDay = string.Empty;
            byDayValue = string.Empty;
            byMonthDay = string.Empty;
            byMonthDayCount = string.Empty;
            byMonth = string.Empty;
            byMonthCount = string.Empty;
            weeklyByDay = string.Empty;
            until = string.Empty;
            untilValue = string.Empty;

            if (ruleArray == null)
            {
                throw new ArgumentNullException(nameof(ruleArray));
            }

            for (int i = 0; i < ruleArray.Length; i++)
            {
                if (ruleArray[i].Equals("COUNT", StringComparison.Ordinal))
                {
                    count = ruleArray[i];
                    recurrenceCount = ruleArray[i + 1];
                }
                else if (ruleArray[i].Equals("DAILY", StringComparison.Ordinal))
                {
                    freq = daily = ruleArray[i];
                }
                else if (ruleArray[i].Equals("WEEKLY", StringComparison.Ordinal))
                {
                    freq = weekly = ruleArray[i];
                }
                else if (ruleArray[i].Equals("INTERVAL", StringComparison.Ordinal))
                {
                    interval = ruleArray[i];
                    intervalCount = ruleArray[i + 1];
                }
                else if (ruleArray[i].Equals("UNTIL", StringComparison.Ordinal))
                {
                    until = ruleArray[i];
                    untilValue = ruleArray[i + 1];
                }
                else if (ruleArray[i].Equals("MONTHLY", StringComparison.Ordinal))
                {
                    freq = monthly = ruleArray[i];
                }
                else if (ruleArray[i].Equals("YEARLY", StringComparison.Ordinal))
                {
                    freq = yearly = ruleArray[i];
                }
                else if (ruleArray[i].Equals("BYSETPOS", StringComparison.Ordinal))
                {
                    bySetPos = ruleArray[i];
                    bySetPosCount = ruleArray[i + 1];
                }
                else if (ruleArray[i].Equals("BYDAY", StringComparison.Ordinal))
                {
                    byDayPosition = i;
                    byDay = ruleArray[i];
                    byDayValue = ruleArray[i + 1];
                }
                else if (ruleArray[i].Equals("BYMONTH", StringComparison.Ordinal))
                {
                    byMonth = ruleArray[i];
                    byMonthCount = ruleArray[i + 1];
                }
                else if (ruleArray[i].Equals("BYMONTHDAY", StringComparison.Ordinal))
                {
                    byMonthDayPosition = i;
                    byMonthDay = ruleArray[i];
                    byMonthDayCount = ruleArray[i + 1];
                }
            }
        }

        public static void FindWeeklyRule(out string[] weeklyRule, out string weeklybyday, out int weeklybydaypos, string[] rule, DayOfWeek startDate)
        {
            weeklybyday = string.Empty;
            weeklybydaypos = 0;
            weeklyRule = rule;
            if (weeklyRule == null)
            {
                throw new ArgumentNullException(nameof(weeklyRule));
            }

            bool isWeeklyRule = false;
            bool hasByDay = false;
            List<string> rules = weeklyRule.ToList();
            for (int i = 0; i < weeklyRule.Length; i++)
            {
                if (weeklyRule[i].Contains("BYDAY"))
                {
                    weeklybyday = weeklyRule[i];
                    weeklybydaypos = i;
                    hasByDay = true;
                    break;
                }

                if (weeklyRule[i].Contains("WEEKLY"))
                {
                    isWeeklyRule = true;
                }
            }

            if (isWeeklyRule && !hasByDay)
            {
                switch (startDate)
                {
                    case DayOfWeek.Sunday:
                        {
                            weeklybyday = "BYDAY=SU";
                            break;
                        }

                    case DayOfWeek.Monday:
                        {
                            weeklybyday = "BYDAY=MO";
                            break;
                        }

                    case DayOfWeek.Tuesday:
                        {
                            weeklybyday = "BYDAY=TU";
                            break;
                        }

                    case DayOfWeek.Wednesday:
                        {
                            weeklybyday = "BYDAY=WE";
                            break;
                        }

                    case DayOfWeek.Thursday:
                        {
                            weeklybyday = "BYDAY=TH";
                            break;
                        }

                    case DayOfWeek.Friday:
                        {
                            weeklybyday = "BYDAY=FR";
                            break;
                        }

                    case DayOfWeek.Saturday:
                        {
                            weeklybyday = "BYDAY=SA";
                            break;
                        }
                }

                rules.Add(weeklybyday);
                weeklybydaypos = rules.Count - 1;
                weeklyRule = rules.ToArray();
            }
#pragma warning restore CA1021 // Avoid out parameters
        }

        #endregion
    }
    public class ScheduleUtils
    {
        internal static DateTime GetWeekFirstDate(DateTime date, double firstDayOfWeek)
        {
            firstDayOfWeek = (firstDayOfWeek - (int)date.DayOfWeek + 7 * (-1)) % 7;
            DateTime dateValue = date.AddDays(firstDayOfWeek);
            return dateValue;
        }

        internal static DateTime FirstDateOfMonth(DateTime date)
        {
            return new DateTime(date.Year, date.Month, 1);
        }

        internal static DateTime LastDateOfMonth(DateTime date)
        {
            DateTime firstDayOfMonth = new DateTime(date.Year, date.Month, 1).AddMonths(1);
            return firstDayOfMonth.AddDays(-1);
        }
    }
}
