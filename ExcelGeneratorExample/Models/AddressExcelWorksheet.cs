using System.Collections.Generic;
using Msi.ExcelGenerator;

namespace ExcelGeneratorExample.Models
{
    public class AddressExcelWorksheet : IExcelWorksheet
    {
        public string Name { get; set; }
        public IEnumerable<object> Models { get; set; }
    }
}
