using System.Collections.Generic;

namespace Msi.ExcelGenerator
{
    public interface IExcelWorksheet
    {
        string Name { get; set; }
        IEnumerable<object> Models { get; set; }
    }
}
