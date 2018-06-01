namespace Msi.ExcelGenerator
{
    public interface IExcelDocument
    {
        IExcelDocument AddWorksheet(IExcelWorksheet worksheet);
        byte[] Generate();
    }
}
