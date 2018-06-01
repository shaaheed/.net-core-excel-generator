namespace Msi.ExcelGenerator
{
    public class ExcelGenerator : IExcelGenerator
    {
        public IExcelDocument NewDocument()
        {
            return new ExcelDocument();
        }
    }
}
