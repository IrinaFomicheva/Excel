using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace XLS
{
    public interface IDataProcessingService
    {
        List<DataItem> GetForecastValues(List<DataItem> actualDataItems, DateTime endDate);
        List<DataItem> GetEstimateValues(List<DataItem> actualValues, DateTime startDate, int indexStartDate);
    }


    public class DataProcessingService : IDataProcessingService
    {
        private List<DataItem> forecastValues;
        private List<DataItem> estimateValues;

        int daysDiff;

        DateTime dateTime;
        decimal value;
        decimal sum;
        decimal rand;

        decimal deltaEstimateValues;
        public DataProcessingService()
        {

        }
        public List<DataItem> GetForecastValues(List<DataItem> actualValues, DateTime endDate)
        {
            daysDiff = (endDate - actualValues[actualValues.Count - 1].Date).Days;

            value = actualValues[actualValues.Count - 1].PtdValue;
            sum = actualValues[actualValues.Count - 1].SumValue;
            dateTime = actualValues[actualValues.Count - 1].Date;
            forecastValues = new List<DataItem>
            {
                new DataItem(dateTime, value, sum)
            };
            var r = new Random();
            for (int i = 0; i < daysDiff; i++)
            {
                rand = Convert.ToDecimal(r.Next(80, 121) / 100.00);
                value = Math.Round((sum * rand / actualValues.Count), 2);
                sum += value;
                dateTime = dateTime.AddDays(1);
                forecastValues.Add(new DataItem(dateTime, value, sum));
            }
            foreach (DataItem i in forecastValues)
            {
                Console.WriteLine("forecastValues  " + i.Date.ToString() + "  " + i.PtdValue + "  " + i.SumValue);
                Console.WriteLine("==================");
            }
            return forecastValues;
        }
        public List<DataItem> GetEstimateValues(List<DataItem> actualValues, DateTime startDate, int indexStartDate)
        {
            dateTime = startDate.AddDays(-1);
            value = 0;
            sum = 0;
            estimateValues = new List<DataItem>
            {
                new DataItem(dateTime, value, sum)
            };

            deltaEstimateValues = actualValues[indexStartDate - 1].SumValue * daysDiff / 10;

            for (int k = indexStartDate; k < actualValues.Count; k++)
            {
                value = actualValues[k].PtdValue + deltaEstimateValues;
                sum += value;
                dateTime = actualValues[k].Date;
                estimateValues.Add(new DataItem(dateTime, value, sum));
            }
            for (int k = 1; k < forecastValues.Count; k++)
            {
                value = forecastValues[k].PtdValue + deltaEstimateValues;
                sum += value;
                dateTime = forecastValues[k].Date;
                estimateValues.Add(new DataItem(dateTime, value, sum));
            }
            foreach (DataItem i in estimateValues)
            {
                Console.WriteLine("estimateValues  " + i.Date.ToString() + "  " + i.PtdValue + "  " + i.SumValue);
                Console.WriteLine("==================");
            }
            return estimateValues;
        }

    }

}
