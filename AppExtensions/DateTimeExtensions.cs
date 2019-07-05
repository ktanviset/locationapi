using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace locationapi.AppExtensions
{
    public static class DateTimeExtensions
    {
        public static DateTime? ToNullableDateTimeValue(this DateTime dateTime)
        {
            DateTime? ndatetime = null;
            DateTime dedatetime = new DateTime();

            if (dateTime != dedatetime)
            {
                ndatetime = dateTime;
            }

            return ndatetime;
        }
    }
}
