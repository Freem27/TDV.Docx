using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TDV.Docx
{
    public class Range<T> where T : IComparable<T>
    {
        public T start;
        public T end;
        public Range(T start, T end)
        {
            this.start = start;
            this.end = end;
        }

        /// <summary>
        /// начало и конец диапазона = valeu
        /// </summary>
        /// <param name="value"></param>
        public Range(T value)
        {
            this.start = value;
            this.end = value;
        }

        bool InRange(T value)
        {
            return start.CompareTo(value) >= 0 && end.CompareTo(value) <= 0;
        }
    }
    static class ListExtentions
    {
        public static int Median(this IEnumerable<int> source)
        {
            int[] temp = source.ToArray();
            Array.Sort(temp);
            int count = temp.Length;
            if (count == 0)
            {
                throw new InvalidOperationException("Список пуст");
            }
            else if (count % 2 == 0)
            {
                int a = temp[count / 2 - 1];
                int b = temp[count / 2];
                return (a + b) / 2;
            }
            else
            {
                return temp[count / 2];
            }
        }
    }

    static class StringExtentions { 
        public static List<int> AllIndexesOf(this string str, string value)
        {
            if (String.IsNullOrEmpty(value))
                throw new ArgumentException("Искомая строка не может быть пустой", "value");
            List<int> indexes = new List<int>();
            for (int index = 0; ; index += value.Length)
            {
                index = str.IndexOf(value, index);
                if (index == -1)
                return indexes;
                indexes.Add(index);
            }
        }
    }

    static class IntExtentions
    {
        public static bool Between(this int source, int start , int end)
        {
            if (start <= source && end >= source)
                return true;
            else
                return false;
        }
    }
    public class Pair<T, U>
    {
        public Pair()
        {
        }

        public Pair(T first, U second)
        {
            this.First = first;
            this.Second = second;
        }

        public T First { get; set; }
        public U Second { get; set; }
    };
}
