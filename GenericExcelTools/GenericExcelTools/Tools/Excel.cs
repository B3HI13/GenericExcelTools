using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Reflection;

namespace GenericExcelTools
{
    public class Excel
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

        public void Export<TDataModel>(List<RqExport<TDataModel>> lstData, string sheetName, string saveFilePath, string countOfRecordsTitle = "تعداد") where TDataModel : new()
        {
            var workBook = new XLWorkbook { RightToLeft = true };
            var workSheet = workBook.Worksheets.Add(sheetName);

            var lstDataModelProperties = typeof(TDataModel).GetProperties().ToList();
            var lstColumnHeaderNames = lstDataModelProperties.ToList();

            #region FillingColumnNamesWithDesignings
            workSheet.Cell(1, 1).Value = "RowNo";

            for (int i = 0; i < lstColumnHeaderNames.Count; i++)
            {
                workSheet.Cell(1, i + 2).Value =
                    (lstColumnHeaderNames[i].GetCustomAttributes(typeof(DisplayAttribute), true).FirstOrDefault(q => q is DisplayAttribute) as DisplayAttribute).Name;
            }

            //The reasone of  + 1 in lastCellColumn is the first column is for "RowNo"
            workSheet.Range(firstCellRow: 1, firstCellColumn: 1, lastCellRow: 1, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Fill.BackgroundColor = XLColor.Blue;
            workSheet.Range(firstCellRow: 1, firstCellColumn: 1, lastCellRow: 1, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Font.FontColor = XLColor.White;
            workSheet.Range(firstCellRow: 1, firstCellColumn: 1, lastCellRow: 1, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Font.Bold = true;
            #endregion

            int insertingBeginRowNumber = 2;
            var defaultDataRecordColor = XLColor.LightBlue;
            foreach (var data in lstData)
            {
                #region FillingDataWithDesignings
                for (int i = 0; i < data.LstData.Count; i++)
                {
                    TDataModel objData = data.LstData[i];

                    //The reasone of  + 1 in lastCellColumn is the first column is for "RowNo"
                    workSheet.Range(firstCellRow: insertingBeginRowNumber, firstCellColumn: 1, lastCellRow: insertingBeginRowNumber, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Fill.BackgroundColor = defaultDataRecordColor;
                    workSheet.Range(firstCellRow: insertingBeginRowNumber, firstCellColumn: 1, lastCellRow: insertingBeginRowNumber, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Font.FontColor = XLColor.Black;
                    workSheet.Range(firstCellRow: insertingBeginRowNumber, firstCellColumn: 1, lastCellRow: insertingBeginRowNumber, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Font.Italic = false;
                    workSheet.Range(firstCellRow: insertingBeginRowNumber, firstCellColumn: 1, lastCellRow: insertingBeginRowNumber, lastCellColumn: lstColumnHeaderNames.Count + 1).Style.Font.Bold = false;

                    if (defaultDataRecordColor == XLColor.LightBlue) defaultDataRecordColor = XLColor.White;
                    else defaultDataRecordColor = XLColor.LightBlue;
                    insertingBeginRowNumber++;
                }

                //The reasone of column: 2 is the column 1 is for "RowNo"
                workSheet.Cell(insertingBeginRowNumber - data.LstData.Count, column: 2).InsertData(data.LstData);

                int counter = 1;
                for (int rowIndex = insertingBeginRowNumber - data.LstData.Count; rowIndex < insertingBeginRowNumber; rowIndex++)
                {
                    workSheet.Cell(rowIndex, column: 1).Value = counter;
                    counter++;
                }
                #endregion
            }

            workSheet.Columns().AdjustToContents();
            workSheet.Rows().AdjustToContents();

            workBook.SaveAs(saveFilePath);
        }
    }
}
