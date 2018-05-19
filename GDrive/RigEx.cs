using System;
using System.Collections.Generic;
using System.Linq;

namespace GDrive
{
    public static class RigEx
    {

        public static string AddTimeStamp(this string s)
        {
            return $"{DateTime.Now:HH:mm:ss tt} : {s}";
        }

        public static void WriteLineColors(string msg, ConsoleColor color)
        {
            var cur = Console.ForegroundColor;
            Console.ForegroundColor = color;
            Console.WriteLine(msg);
            Console.ForegroundColor = cur;
        }

        public static TimeSpan ToTimeSpan(int minutes)
        {
            float d = minutes / 1440f;
            float h = (d - (int)d) * 24;
            float m = minutes % 1440f % 60f;
            return new TimeSpan((int)d, (int)h, (int)m, 0);
        }

        public static int LaterToColum(this string s)
        {
            return int.Parse(s.ToUpper(), System.Globalization.NumberStyles.HexNumber) - 10;
        }

        public static void WriteLog(string message)
        {
            Console.WriteLine($"{DateTime.Now:HH:mm:ss tt} : {message}");
        }

        public static IEnumerable<I> As<T, I>(this List<T> list) where T : I
        {
            return list.Cast<I>();
        }

      
    }
}
