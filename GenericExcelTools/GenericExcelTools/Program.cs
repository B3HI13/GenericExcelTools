using DocumentFormat.OpenXml.Office2010.ExcelAc;
using System;
using System.Collections.Generic;
using System.IO;

namespace GenericExcelTools
{
    class Program
    {
        static void Main(string[] args)
        {
            Console.WriteLine("To read data from excel, please press 0. \nTo export data to excel, please press 1.");
            string requestNumber = Console.ReadLine();
            if (requestNumber == "0")
                Read();
            else if ((requestNumber == "1"))
                Export();
            else
                throw new ArgumentOutOfRangeException("Please enter valid input");
        }

        private static void Read()
        {
            Excel objExcel = new Excel();
            RsRead<TestModel> objRsRead = objExcel.Read<TestModel>(new RqRead
            {
                ExcelFilePath = Directory.GetCurrentDirectory() + "\\Resources\\ReadTest.xlsx"
            });

            objRsRead.LstData.ForEach(q => { Console.WriteLine($"{nameof(q.Id)}: {q.Id}, {nameof(q.Name)}: {q.Name}\n"); });
        }

        private static void Export()
        {
            Excel objExcel = new Excel();

            var lstData = new List<RqExport<TestModel>>
            {
                new RqExport<TestModel>
                {
                    LstData = new List<TestModel>
                    {
                        new TestModel{ Id = "1", Name = "TestExport1"},
                        new TestModel{ Id = "2", Name = "TestExport2"},
                        new TestModel{ Id = "3", Name = "TestExport3"},
                    }
                }
            };

            objExcel.Export(
                lstData,
                sheetName: "TestExport",
                saveFilePath: Directory.GetCurrentDirectory() + $"\\Resources\\TestExport{new Random().Next(0, int.MaxValue)}.xlsx");

            Console.WriteLine("Done");
        }
    }
}
