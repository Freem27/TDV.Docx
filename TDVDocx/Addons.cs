using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TDV.Docx
{
    public class Range<T> where T: IComparable<T>
    {
        public T start;
        public T end;
        public Range(T start, T end)
        {
            this.start=start;
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
            return start.CompareTo(value)>=0 && end.CompareTo(value)<=0;
        }

        
    }
}
