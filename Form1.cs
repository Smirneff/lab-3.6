using lab3.Variant4;
using lab3.Variant8;
using NPOI.SS.Formula.Functions;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.Linq;
using System.Windows.Forms;
using ZedGraph;
namespace lab3
{
    public partial class Stats : Form
    {
        GraphPane pane;
        public Stats()
        {
            InitializeComponent();
            pane = zedGraphControl1.GraphPane;
        }
        private void button1_Click(object sender, EventArgs e)
        {
            // Очистим список кривых на тот случай, если до этого сигналы уже были нарисованы
            pane.CurveList.Clear();
            GDPDataProcessor dataProcessor = new GDPDataProcessor();

            // Загрузка данных из файла Excel
            dataProcessor.LoadDataFromExcel("../../Variant4/variant4.xlsx");

            // Вывод данных на экран
            dataProcessor.DisplayData();

            // Создание объекта PointPairList для ВВП и заполнение его значениями из массива данных
            PointPairList gdpPointPairList = new PointPairList();
            for (int i = 0; i < dataProcessor.GDPDataList.Count; i++)
            {
                gdpPointPairList.Add(new XDate(dataProcessor.GDPDataList[i].Year, 1, 1), dataProcessor.GDPDataList[i].GDP);
            }

            // Добавление объекта PointPairList для ВВП в CurveList графической панели
            pane.CurveList.Add(new LineItem("ВВП", gdpPointPairList, Color.Blue, SymbolType.Circle));

            // Создание объекта PointPairList для ВНП и заполнение его значениями из массива данных
            PointPairList gnpPointPairList = new PointPairList();
            for (int i = 0; i < dataProcessor.GDPDataList.Count; i++)
            {
                gnpPointPairList.Add(new XDate(dataProcessor.GDPDataList[i].Year, 1, 1), dataProcessor.GDPDataList[i].GNP);
            }

            // Добавление объекта PointPairList для ВНП в CurveList графической панели
            pane.CurveList.Add(new LineItem("ВНП", gnpPointPairList, Color.Red, SymbolType.Circle));

            // Пример изменения масштаба по оси, где откладываются годы
            pane.XAxis.Scale.Min = new XDate(dataProcessor.GDPDataList[0].Year, 1, 1);
            pane.XAxis.Scale.Max = new XDate(dataProcessor.GDPDataList[dataProcessor.GDPDataList.Count - 1].Year, 1, 1);

            pane.XAxis.Type = AxisType.Date;
            pane.XAxis.Scale.Format = "yyyy";
            pane.XAxis.Scale.MajorStep = 1;

            // Обновление отображения графика
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();

            dataProcessor.CalculateGrowthRate();
        }

        private void button2_Click(object sender, EventArgs e)
        {
            SalaryDataProcessor dataProcessor = new SalaryDataProcessor();

// Загрузка данных из файла Excel
dataProcessor.LoadDataFromExcel("../../Variant8/SalaryData.xlsx");

            // Вывод данных на экран
            dataProcessor.DisplayData();

            // Очистим список кривых на тот случай, если до этого графики уже были нарисованы
            pane.CurveList.Clear();

            // Создание объекта PointPairList для зарплаты мужчин и заполнение его значениями из массива данных
            PointPairList maleSalaryPointPairList = new PointPairList();
            List<int> years = dataProcessor.Years;
            List<double> maleSalaries = dataProcessor.MaleSalaries;
            for (int i = 0; i < maleSalaries.Count; i++)
            {
                maleSalaryPointPairList.Add(new XDate(years[i], 1, 1), maleSalaries[i]);
            }

            // Добавление объекта PointPairList для зарплаты мужчин в CurveList графической панели
            pane.CurveList.Add(new LineItem("Зарплата мужчин", maleSalaryPointPairList, Color.Blue, SymbolType.Circle));

            // Создание объекта PointPairList для зарплаты женщин и заполнение его значениями из массива данных
            PointPairList femaleSalaryPointPairList = new PointPairList();
            List<double> femaleSalaries = dataProcessor.FemaleSalaries;
            for (int i = 0; i < femaleSalaries.Count; i++)
            {
                femaleSalaryPointPairList.Add(new XDate(years[i], 1, 1), femaleSalaries[i]);
            }

            // Добавление объекта PointPairList для зарплаты женщин в CurveList графической панели
            pane.CurveList.Add(new LineItem("Зарплата женщин", femaleSalaryPointPairList, Color.Red, SymbolType.Circle));

            // Пример изменения масштаба по оси, где откладываются годы
            pane.XAxis.Scale.Min = new XDate(years[0] , 1, 1);
            pane.XAxis.Scale.Max = new XDate(years[years.Count - 1] , 1, 1);

            pane.XAxis.Type = AxisType.Date;
            pane.XAxis.Scale.Format = "yyyy";
            pane.XAxis.Scale.MajorStep = 1;

            // Обновление отображения графика
            zedGraphControl1.AxisChange();
            zedGraphControl1.Invalidate();

            Console.WriteLine(dataProcessor.CalculateMaleSalaryGrowthRate());
            Console.WriteLine(dataProcessor.CalculateFemaleSalaryGrowthRate());
        }
    }
}
