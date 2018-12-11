using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelFunctions
{
    public class Date
    {

        /// <summary>
        /// Returns year from date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eYear(DateTime date)
        {
            return date.Year;
        }
        /// <summary>
        /// Returns year from date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eYear(int date)
        {
            DateTime start = new DateTime(1900, 1, 1);
            start = start.AddDays(date).AddDays(-2);
            int end = start.Year;
            return end;
        }
        /// <summary>
        /// Returns month from date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eMonth(DateTime date)
        {
            return date.Month;
        }

        /// <summary>
        /// Returns month from date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eMonth(int date)
        {
            DateTime start = new DateTime(1900, 1, 1);
            start = start.AddDays(date).AddDays(-2);
            int end = start.Month;
            return end;
        }

        /// <summary>
        /// Returns day of a month from date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eDay(DateTime date)
        {
            return date.Day;
        }

        /// <summary>
        /// Returns day of a month from date
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eDay(int date)
        {
            DateTime start = new DateTime(1900, 1, 1);
            start = start.AddDays(date).AddDays(-2);
            int end = start.Day;
            return end;
        }

        /// <summary>
        /// Returns day of a week from date (0- Sunday, 1-Monday...)
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eWeekday(DateTime date)
        {
            return Convert.ToInt32(date.DayOfWeek);
        }

        /// <summary>
        /// Returns day of a week from date (0- Sunday, 1-Monday...)
        /// </summary>
        /// <param name="date"></param>
        /// <returns></returns>
        public int eWeekday(int date)
        {
            DateTime start = new DateTime(1900, 1, 1);
            start = start.AddDays(date).AddDays(-2);
            int end = Convert.ToInt32(start.DayOfWeek);
            return end;
        }

        /// <summary>
        /// Returns the current date
        /// </summary>
        /// <returns></returns>
        public DateTime eToday()
        {
            return DateTime.Today;
        }

        public enum spans { y, m, d };
        /// <summary>
        /// Calculates the number of days, months, or years between two dates
        /// </summary>
        /// <param name="date1">less than date2</param>
        /// <param name="date2">greater than date1</param>
        /// <param name="span">y - years, m- months, d- days</param>
        /// <returns></returns>
        public int eDateDif(DateTime date1, DateTime date2, spans span)
        {
            int dif;
            DateTime start1 = date1;
            DateTime start2 = date2;
            if (span == spans.y)
                if (start1.Year < start2.Year)
                    dif = Convert.ToInt32((start2.Year - start1.Year - 1) + (((start2.Month > start1.Month) || ((start2.Month == start1.Month) && (start2.Day >= start1.Day))) ? 1 : 0));
                else if (start1.Year == start2.Year)
                    dif = 0;
                else
                    dif = -Convert.ToInt32((start1.Year - start2.Year - 1) + (((start1.Month > start2.Month) || ((start1.Month == start2.Month) && (start1.Day >= start2.Day))) ? 1 : 0));
            else if (span == spans.m)
                if (start1.Month <= start2.Month)
                    dif = Convert.ToInt32((start2.Year - start1.Year) * 12 + start2.Month - start1.Month + (start2.Day >= start1.Day ? 0 : -1));
                else
                    dif = -Convert.ToInt32((start1.Year - start2.Year) * 12 + start1.Month - start2.Month + (start1.Day >= start2.Day ? 0 : -1));
            else if (span == spans.d)
                if (start1.Day <= start2.Day)
                    dif = Convert.ToInt32((date2 - date1).TotalDays);
                else
                    dif = -Convert.ToInt32((date1 - date2).TotalDays);
            else throw new ArgumentException("Wrong span");
            return dif;
        }

        /// <summary>
        /// Calculates the number of days, months, or years between two dates
        /// </summary>
        /// <param name="date1">less than date2</param>
        /// <param name="date2">greater than date1</param>
        /// <param name="span">y - years, m- months, d- days</param>
        /// <returns></returns>
        public int eDateDif(int date1, int date2, spans span)
        {
            DateTime start1 = new DateTime(1900, 1, 1);
            start1 = start1.AddDays(date1);
            DateTime start2 = new DateTime(1900, 1, 1);
            start2 = start2.AddDays(date2);

            int dif;

            if (span == spans.y)
                if (start1.Year < start2.Year)
                    dif = Convert.ToInt32((start2.Year - start1.Year - 1) + (((start2.Month > start1.Month) || ((start2.Month == start1.Month) && (start2.Day >= start1.Day))) ? 1 : 0));
                else if (start1.Year == start2.Year)
                    dif = 0;
                else
                    dif = -Convert.ToInt32((start1.Year - start2.Year - 1) + (((start1.Month > start2.Month) || ((start1.Month == start2.Month) && (start1.Day >= start2.Day))) ? 1 : 0));
            else if (span == spans.m)
                if (start1.Month <= start2.Month)
                    dif = Convert.ToInt32((start2.Year - start1.Year) * 12 + start2.Month - start1.Month + (start2.Day >= start1.Day ? 0 : -1));
                else
                    dif = -Convert.ToInt32((start1.Year - start2.Year) * 12 + start1.Month - start2.Month + (start1.Day >= start2.Day ? 0 : -1));
            else if (span == spans.d)
                if (start1.Day <= start2.Day)
                    dif = date2 - date1;
                else
                    dif = -(date1 - date2);
            else throw new ArgumentException("Wrong span");
            return dif;
        }

        /// <summary>
        /// Adds a specified number of months to a date and returns date
        /// </summary>
        /// <param name="date"></param>
        /// <param name="monthsNo"></param>
        /// <returns></returns>
        public DateTime eEDate(int date, int monthsNo)
        {
            DateTime start = new DateTime(1900, 1, 1);
            start = start.AddDays(date).AddDays(-2);

            return start.AddMonths(monthsNo);
        }

        /// <summary>
        /// Adds a specified number of months to a date, returns date
        /// </summary>
        /// <param name="date"></param>
        /// <param name="monthsNo"></param>
        /// <returns></returns>
        public DateTime eEDate(DateTime date, int monthsNo)
        {
            return date.AddMonths(monthsNo);
        }

        /// <summary>
        /// Calculates the last day of the month after adding a specified number of months to a date, returns date
        /// </summary>
        /// <param name="date"></param>
        /// <param name="monthsNo"></param>
        /// <returns></returns>
        public DateTime eEomonth(int date, int monthsNo)
        {
            DateTime start = new DateTime(1900, 1, 1);
            start = start.AddDays(date).AddDays(-2);

            var firstDayOfMonth = new DateTime(start.Year, start.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(monthsNo + 1).AddDays(-1);

            return lastDayOfMonth;
        }

        /// <summary>
        /// Calculates the last day of the month after adding a specified number of months to a date, returns date
        /// </summary>
        /// <param name="date"></param>
        /// <param name="monthsNo"></param>
        /// <returns></returns>
        public DateTime eEomonth(DateTime date, int monthsNo)
        {
            var firstDayOfMonth = new DateTime(date.Year, date.Month, 1);
            var lastDayOfMonth = firstDayOfMonth.AddMonths(monthsNo + 1).AddDays(-1);

            return lastDayOfMonth;
        }
    }
}
