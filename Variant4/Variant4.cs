using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;

namespace lab3.Variant4
{
    public class GDPDataProcessor
    {
        public List<GDPData> GDPDataList;

        public GDPDataProcessor()
        {
            GDPDataList = new List<GDPData>();
        }

        public void GenerateData()
        {
            Random random = new Random();

            for (int i = 1; i <= 15; i++)
            {
                double gdp = random.Next(1000, 10000);
                double gnp = random.Next(1000, 10000);

                GDPData data = new GDPData(i, gdp, gnp);
                GDPDataList.Add(data);
            }
        }

        public void SaveDataToExcel(string filePath)
        {
            IWorkbook workbook = new XSSFWorkbook();
            ISheet worksheet = workbook.CreateSheet("GDPData");

            // Заголовки столбцов
            IRow headerRow = worksheet.CreateRow(0);
            headerRow.CreateCell(0).SetCellValue("Year");
            headerRow.CreateCell(1).SetCellValue("GDP");
            headerRow.CreateCell(2).SetCellValue("GNP");

            int row = 1;
            foreach (var data in GDPDataList)
            {
                IRow dataRow = worksheet.CreateRow(row);
                dataRow.CreateCell(0).SetCellValue(data.Year);
                dataRow.CreateCell(1).SetCellValue(data.GDP);
                dataRow.CreateCell(2).SetCellValue(data.GNP);

                row++;
            }

            using (FileStream fileStream = new FileStream(filePath, FileMode.Create, FileAccess.Write))
            {
                workbook.Write(fileStream, false);
            }
        }

        public void LoadDataFromExcel(string filePath)
        {
            if (File.Exists(filePath))
            {
                GDPDataList.Clear();

                using (FileStream fileStream = new FileStream(filePath, FileMode.Open, FileAccess.Read))
                {
                    IWorkbook workbook = new XSSFWorkbook(fileStream);
                    ISheet worksheet = workbook.GetSheetAt(0);

                    for (int row = 1; row <= worksheet.LastRowNum; row++)
                    {
                        IRow dataRow = worksheet.GetRow(row);

                        int year = (int)dataRow.GetCell(0).NumericCellValue;
                        double gdp = dataRow.GetCell(1).NumericCellValue;
                        double gnp = dataRow.GetCell(2).NumericCellValue;

                        GDPData data = new GDPData(year, gdp, gnp);
                        GDPDataList.Add(data);
                    }
                }
            }
            else
            {
                Console.WriteLine("Файл не найден.");
            }
        }

        public void DisplayData()
        {
            Console.WriteLine("Данные о ВВП и ВНП России за последние 15 лет:");
            Console.WriteLine("Year\tGDP\tGNP");
            foreach (var data in GDPDataList)
            {
                Console.WriteLine($"{data.Year}\t{data.GDP}\t{data.GNP}");
            }
        }

        public void CalculateGrowthRate()
        {
            Console.WriteLine("\nВычисление процента роста/падения ВВП и ВНП:");

            for (int i = 1; i < GDPDataList.Count; i++)
            {
                double currentGDP = GDPDataList[i].GDP;
                double previousGDP = GDPDataList[i - 1].GDP;
                double gdpGrowthRate = ((currentGDP - previousGDP) / previousGDP) * 100;

                double currentGNP = GDPDataList[i].GNP;
                double previousGNP = GDPDataList[i - 1].GNP;
                double gnpGrowthRate = ((currentGNP - previousGNP) / previousGNP) * 100;

                Console.WriteLine($"Год {GDPDataList[i].Year}:");
                Console.WriteLine($"Процент роста/падения ВВП: {gdpGrowthRate}%");
                Console.WriteLine($"Процент роста/падения ВНП: {gnpGrowthRate}%");
            }
        }
    }

    public class GDPData
    {
        public int Year { get; set; }
        public double GDP { get; set; }
        public double GNP { get; set; }

        public GDPData(int year, double gdp, double gnp)
        {
            Year = year;
            GDP = gdp;
            GNP = gnp;
        }
    }
}
