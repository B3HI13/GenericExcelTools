using System;
using System.IO;

namespace GenericExcelTools
{
    class Program
    {
        static void Main(string[] args)
        {
            ExcelReader objExcelReader = new ExcelReader();
            RsRead<TestModel> objRsRead = objExcelReader.Read<TestModel>(new RqRead
            {
                ExcelFilePath = Directory.GetCurrentDirectory() + "\\Resources\\test.xlsx"
            });

            objRsRead.LstData.ForEach(q => { Console.WriteLine($"{nameof(q.Id)}: {q.Id}, {nameof(q.Name)}: {q.Name}\n"); });
        }
    }
}
