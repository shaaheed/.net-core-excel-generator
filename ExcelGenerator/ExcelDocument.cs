using OfficeOpenXml;
using System.Collections.Generic;
using System.IO;
using System.Linq;

namespace Msi.ExcelGenerator
{
    public class ExcelDocument : IExcelDocument
    {

        private List<IExcelWorksheet> _worksheets = new List<IExcelWorksheet>();

        public IExcelDocument AddWorksheet(IExcelWorksheet worksheet)
        {
            _worksheets.Add(worksheet);
            return this;
        }

        public byte[] Generate()
        {
            using (var stream = new MemoryStream())
            {
                using (var package = new ExcelPackage(stream))
                {
                    foreach (var worksheet in _worksheets)
                    {
                        ExcelWorksheet sheet = package.Workbook.Worksheets.Add(worksheet.Name);
                        var firstModel = worksheet.Models.FirstOrDefault();
                        var properties = firstModel.GetType().GetProperties();
                        for (int i = 0; i < properties.Length; i++)
                        {
                            sheet.Cells[1, i + 1].Value = properties[i].Name;
                        }
                        int row = 2;
                        foreach (var model in worksheet.Models)
                        {
                            for (int j = 0; j < properties.Length; j++)
                            {
                                sheet.Cells[row, j + 1].Value = properties[j].GetValue(model);
                            }
                            row++;
                        }
                    }
                    package.Save();
                }
                return stream.ToArray();
            }
        }
    }
}
