using System.Collections.Generic;

namespace GenericExcelTools
{
    public class RqExport<TDataModel>
    {
        public List<TDataModel> LstData { get; set; }

        public RqExport()
        {
            LstData = new List<TDataModel>();
        }
    }
}
