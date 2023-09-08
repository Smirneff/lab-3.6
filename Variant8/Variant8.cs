using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace lab3.Variant8
{
    public class SalaryDataProcessor
    {
        public List<int> Years { get; private set; }
        public List<double> MaleSalaries { get; private set; }
        public List<double> FemaleSalaries { get; private set; }

        public SalaryDataProcessor()
        {
            Years = new List<int>();
            MaleSalaries = new List<double>();
            FemaleSalaries = new List<double>();
        }

        public void GenerateData(int numberOfYears)
        {
            // Генерация случайных данных о медианной заработной плате
            Random random = new Random();

            for (int i = 0; i < numberOfYears; i++)
            {
                int year = DateTime.Now.Year - numberOfYears + i + 1;
                double maleSalary = random.NextDouble() * 50000 + 30000;  // Пример случайных данных
                double femaleSalary = random.NextDouble() * 50000 + 30000;  // Пример случайных данных

                Years.Add(year);
                MaleSalaries.Add(maleSalary);
                FemaleSalaries.Add(femaleSalary);
            }
        }

        public void SaveDataToExcel(string filePath)
        {
            // Сохранение данных в файл Excel
            IWorkbook workbook = new XSSFWorkbook();
            ISheet sheet = workbook.CreateSheet("SalaryData");

            // Создание заголовков столбцов
            IRow headerRow = sheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Year");
            headerRow.CreateCell(1).SetCellValue("Male Salary");
            headerRow.CreateCell(2).SetCellValue("Female Salary");

            // Запись данных
            for (int i = 0; i < Years.Count; i++)
            {
                IRow dataRow = sheet.CreateRow(i + 1);
                dataRow.CreateCell(0).SetCellValue(Years[i]);
                dataRow.CreateCell(1).SetCellValue(MaleSalaries[i]);
                dataRow.CreateCell(2).SetCellValue(FemaleSalaries[i]);
            }

            // Сохранение файла
            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream,false);
            }
        }

        public void LoadDataFromExcel(string filePath)
        {
            Years.Clear();
            MaleSalaries.Clear();
            FemaleSalaries.Clear();

            using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
            {
                IWorkbook workbook = new XSSFWorkbook(fileStream);
                ISheet sheet = workbook.GetSheetAt(0);

                // Чтение данных
                for (int i = 1; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row != null)
                    {
                        int year = (int)row.GetCell(0).NumericCellValue;
                        double maleSalary = row.GetCell(1).NumericCellValue;
                        double femaleSalary = row.GetCell(2).NumericCellValue;

                        Years.Add(year);
                        MaleSalaries.Add(maleSalary);
                        FemaleSalaries.Add(femaleSalary);
                    }
                }
            }
        }

        public void DisplayData()
        {
            // Вывод данных на экран
            Console.WriteLine("Year\tMale Salary\tFemale Salary");

            for (int i = 0; i < Years.Count; i++)
            {
                Console.WriteLine($"{Years[i]}\t{MaleSalaries[i]}\t{FemaleSalaries[i]}");
            }
        }

        public double CalculateSalaryGrowthRate(List<double> salaries)
        {
            // Вычисление процента роста зарплаты
            if (salaries.Count < 2)
            {
                throw new ArgumentException("Data should contain at least two values for calculating the growth rate.");
            }

            double firstValue = salaries[0];
            double lastValue = salaries[salaries.Count - 1];

            double growthRate = (lastValue - firstValue) / firstValue * 100;

            return growthRate;
        }

        public double CalculateMaleSalaryGrowthRate()
        {
            return CalculateSalaryGrowthRate(MaleSalaries);
        }

        public double CalculateFemaleSalaryGrowthRate()
        {
            return CalculateSalaryGrowthRate(FemaleSalaries);
        }
    }
}
