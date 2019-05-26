using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using xl = Microsoft.Office.Interop.Excel;
using Microsoft.VisualStudio.Tools.Applications.Runtime;
using System.Windows.Forms.DataVisualization.Charting;

namespace XLS
{
    public partial class Form1 : Form
    {
        private List<DataItem> actualValues;
        private List<DataItem> forecastValues;
        private List<DataItem> estimateValues;

        string stDateTextBox;
        string enDateTextBox;

        DateTime startDate;
        DateTime endDate;

        int indexStartDate = 0;
        int indexEndDate;

        ExcelDataProvider file;
        DataProcessingService dataProcess;

        public Form1()
        {
            InitializeComponent();
        }

        private void Button1_Click(object sender, EventArgs e)
        {
            if (actualValues.Count == 0) MessageBox.Show("Need to click button 'Read Excel File' and point the file");
            if (stDateTextBox == "" || enDateTextBox == "") MessageBox.Show("Need to point Start Date and End Date.\nPlease, enter Dates in fortam 'dd.mm.yyyy'.");

            stDateTextBox = textBox1.Text;
            enDateTextBox = textBox2.Text;
            startDate = DateTime.Parse(stDateTextBox);
            endDate = DateTime.Parse(enDateTextBox);
            if (startDate < actualValues[0].Date || startDate > actualValues[actualValues.Count - 1].Date)
            {
                MessageBox.Show("Need to point Start Date at interval from " + String.Format("{0: dd.MM.yyyy}", actualValues[1].Date) + " to " + String.Format("{0: dd.MM.yyyy}", actualValues[actualValues.Count - 1].Date));
            }
            if (startDate > endDate)
            {
                MessageBox.Show("End Date must be later than Start Date!");
            }

            dataProcess = new DataProcessingService(); // to create dataProcess

            indexEndDate = actualValues.Count - 1;
            for (int i = 0; i < actualValues.Count; i++)
            {
                if (actualValues[i].Date == startDate) //to define indexes in list for: startDate и endDate
                {
                    indexStartDate = i;
                }
                if (actualValues[i].Date == endDate)
                {
                    indexEndDate = i;
                }
            }

            if (indexEndDate == actualValues.Count - 1 & endDate != actualValues[actualValues.Count - 1].Date)
            {
                forecastValues = dataProcess.GetForecastValues(actualValues, endDate);    //to create list forecastValues 
            }
            estimateValues = dataProcess.GetEstimateValues(actualValues, startDate, indexStartDate);   //to create list estimateValues

            chart1.Series.Clear();
            Axis ax = new Axis
            {
                Title = "time [d]"
            };
            chart1.ChartAreas[0].AxisX = ax;
            Axis ay = new Axis
            {
                Title = "Costs [$]"
            };
            chart1.ChartAreas[0].AxisY = ay;
            DrawValues("ActualValues", Color.Red, actualValues);
            DrawValues("ForecastValues", Color.Blue, forecastValues);
            DrawValues("EstimateValues", Color.Green, estimateValues);
            Console.WriteLine("=======================");

        }
        private void DrawValues(string name, Color color, List<DataItem> valueues)
        {
            chart1.Series.Add(name);
            chart1.Series[name].Color = color;
            chart1.Series[name].ChartType = SeriesChartType.Spline;
            if (radioButton1.Checked) chart1.Series[name].ChartType = SeriesChartType.Spline;
            else if (radioButton2.Checked) chart1.Series[name].ChartType = SeriesChartType.Line;
            else if (radioButton3.Checked) chart1.Series[name].ChartType = SeriesChartType.Column;
            else if (radioButton4.Checked) chart1.Series[name].ChartType = SeriesChartType.SplineArea;
            else if (radioButton5.Checked) chart1.Series[name].ChartType = SeriesChartType.Point;
            else if (radioButton6.Checked) chart1.Series[name].ChartType = SeriesChartType.StepLine;
            foreach (var valueue in valueues)
            {
                chart1.Series[name].Points.AddXY(valueue.Date, valueue.SumValue);
            }
        }

        private void Button2_Click(object sender, EventArgs e) // button "Write Forecast"
        {
            file.WriteForecast(forecastValues);
        }

        private void Button3_Click(object sender, EventArgs e)  // button "Read Excel File"
        {
            if (openFileDialog1.ShowDialog() == DialogResult.OK)
            {
                file = new ExcelDataProvider(openFileDialog1.FileName);
                actualValues = file.GetActualValues();
            }
        }
    }
}
