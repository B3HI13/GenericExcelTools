using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;

namespace GenericExcelTools
{
    public class ExcelReader
    {
        public RsRead<TOutputModel> Read<TOutputModel>(RqRead objIn) where TOutputModel : new()
        {
            RsRead<TOutputModel> objOutput = new RsRead<TOutputModel>();

            using (XLWorkbook xlWorkBook = new XLWorkbook(objIn.ExcelFilePath))
            {
                List<IXLRangeRow> lstRows = xlWorkBook
                    .Worksheet(1)
                    .RangeUsed()
                    .RowsUsed()
                    .ToList();

                List<IXLCell> lstWantingCells = lstRows[0].Cells()
                    .Where(q =>
                           objIn.LstWantingColumnNames == null
                           || !objIn.LstWantingColumnNames.Any()
                           || objIn.LstWantingColumnNames.Contains(q.Value))
                    .ToList();

                List<PropertyInfo> lstOutputModelProperties = typeof(TOutputModel).GetProperties().ToList();
                Dictionary<string, int> propertyCellNumberDictionary = new Dictionary<string, int>();
                foreach (PropertyInfo objPropertyInfo in lstOutputModelProperties)
                    foreach (IXLCell objCell in lstWantingCells)
                        if (objCell.Value.ToString() == typeof(TOutputModel).GetPropertyDisplayName(objPropertyInfo.Name))
                            propertyCellNumberDictionary.Add(objPropertyInfo.Name, objCell.Address.ColumnNumber);

                for (int i = 1; i < lstRows.Count; i++)
                {
                    TOutputModel objTOutputModel = new TOutputModel();

                    foreach (PropertyInfo objPropertyInfo in lstOutputModelProperties)
                    {
                        if (objPropertyInfo == null || !objPropertyInfo.CanWrite) continue;

                        string value = lstRows[i].Cell(columnNumber: propertyCellNumberDictionary.Single(q => q.Key == objPropertyInfo.Name).Value).Value.ToString();
                        if (value.HaveNumbers())
                            value.ToStandardNumbers();

                        objPropertyInfo.SetValue(objTOutputModel, value, index: null);
                    }

                    objOutput.LstData.Add(objTOutputModel);
                }
            }

            return objOutput;
        }
    }
}
