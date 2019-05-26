using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLS
{

    public class DataItem
    {
        public DateTime Date { get; set; }
        public decimal PtdValue { get; set; }
        public decimal SumValue { get; set; }
        public DataItem()
        {
        }

        public DataItem(DateTime dt, decimal val, decimal sum)
        {
            Date = dt;
            PtdValue = val;
            SumValue = sum;
        }
    }
}
