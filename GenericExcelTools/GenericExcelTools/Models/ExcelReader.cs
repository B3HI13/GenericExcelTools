using System.Collections.Generic;

namespace GenericExcelTools
{
    public class RqRead
    {
        public string ExcelFilePath { get; set; }

        public List<string> LstWantingColumnNames { get; set; }
    }

    public class RsRead<TOutputModel>
    {
        public List<TOutputModel> LstData { get; set; }
        public RsRead()
        {
            LstData = new List<TOutputModel>();
        }
    }
}
