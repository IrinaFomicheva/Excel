using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using _excel = Microsoft.Office.Interop.Excel;

namespace XLS
{

    public interface IDataProvider
    {
        List<DataItem> GetActualValues();
        void WriteForecast(List<DataItem> forecastValues);
    }

    public class ExcelDataProvider : IDataProvider
    {
        private readonly string _fileName;
        int rowCount = 0;
        private List<DataItem> actualValues = new List<DataItem>();

        public ExcelDataProvider(string fileName)
        {
            _fileName = fileName;
        }
        public List<DataItem> GetActualValues()
        {
            _excel.Application excel = new _excel.Application();
            _excel.Workbook workBook = excel.Workbooks.Open(_fileName);
            _excel.Worksheet workSheet = excel.ActiveSheet as _excel.Worksheet;
            var usedRange = workSheet.UsedRange;
            rowCount = usedRange.Rows.Count;
            actualValues = new List<DataItem>();
            DateTime dateTime;

            dateTime = workSheet.Cells[1, 1].Value;
            dateTime = dateTime.AddDays(-1);
            actualValues.Add(new DataItem(dateTime, 0, 0));
            decimal sum;
            sum = 0;
            for (int i = 1; i <= rowCount; i++)
            {
                dateTime = workSheet.Cells[i, 1].Value;
                decimal value = Convert.ToDecimal(workSheet.Cells[i, 2].Value);
                sum += value;
                actualValues.Add(new DataItem(dateTime, value, sum));
            }
            foreach (DataItem i in actualValues)
            {
                Console.WriteLine("actualValues  " + i.Date.ToString() + "  " + i.PtdValue + "  " + i.SumValue);
                Console.WriteLine("==================");
            }

            workBook.Close(true, Type.Missing, Type.Missing);
            excel.Quit();
            return actualValues;
        }
        public void WriteForecast(List<DataItem> forecastValues)
        {
            _excel.Application excel = new _excel.Application();
            _excel.Workbook workBook = excel.Workbooks.Add();
            _excel.Worksheet workSheet = excel.ActiveSheet as _excel.Worksheet;
            workSheet.Cells.ClearContents();

            rowCount = 0;
            for (int k = 1; k < forecastValues.Count; k++)
            {
                rowCount += 1;
                workSheet.Cells[rowCount, 1].Value = forecastValues[k].Date;
                workSheet.Cells[rowCount, 2].Value = forecastValues[k].PtdValue;
            }
            workSheet.Cells[1, 1].EntireColumn.Autofit();
            workBook.Close(true);
            excel.Quit();
        }
    }

}
