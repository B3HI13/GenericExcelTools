
using System.ComponentModel.DataAnnotations;

namespace GenericExcelTools
{
    public class TestModel
    {
        [Display(Name = "id")]
        public string Id { get; set; }

        [Display(Name = "name")]
        public string Name { get; set; }
    }
}
