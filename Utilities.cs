using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TelevendFilter
{
    class Utilities
    {
        public static string FormSQLDateFormat(DateTime? date, bool beginning)
        {
            string output = "";

            output += date.Value.Year + "-";
            output += CheckMonth(date.Value.Month) + "-";
            output += CheckDay(date.Value.Day) + " ";

            if (beginning)
            {
                output += "00:00:00";
            }
            else
            {
                output += "23:59:59";
            }
             
            return output;
        }

        static string CheckMonth(int month)
        {
            if(month < 10)
            {
                return "0" + month.ToString();
            }
            else
            {
                return month.ToString();
            }
        }
        
        static string CheckDay(int day)
        {
            if(day < 10)
            {
                return "0" + day.ToString();
            }
            else
            {
                return day.ToString();
            }
        }
    }
}
