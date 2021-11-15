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
}
