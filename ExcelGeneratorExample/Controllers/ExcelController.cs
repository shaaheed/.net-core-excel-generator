using System.Collections.Generic;
using ExcelGeneratorExample.Models;
using Microsoft.AspNetCore.Mvc;
using Msi.ExcelGenerator;

namespace ExcelGeneratorExample.Controllers
{
    [Route("api")]
    public class ExcelController : Controller
    {
        private readonly IExcelGenerator _excelGenerator;

        public ExcelController(IExcelGenerator excelGenerator)
        {
            _excelGenerator = excelGenerator;
        }

        // POST
        [HttpGet("excel")]
        public IActionResult Post()
        {

            UserExcelWorksheet userExcelWorksheet = new UserExcelWorksheet
            {
                Name = "Users",
                Models = new List<object>()
                {
                    new User{ Name = "A", Mobile = "1" },
                    new User{ Name = "B", Mobile = "2" },
                    new User{ Name = "C", Mobile = "3" }
                }
            };

            AddressExcelWorksheet addressExcelWorksheet = new AddressExcelWorksheet
            {
                Name = "Addressess",
                Models = new List<object>()
                {
                    new Address{ Street = "A", City = "B" },
                    new Address{ Street = "Y", City = "Z" },
                    new Address{ Street = "AY", City = "BZ" }
                }
            };

            var excel = _excelGenerator.NewDocument();
            excel.AddWorksheet(userExcelWorksheet);
            excel.AddWorksheet(addressExcelWorksheet);

            byte[] result = excel.Generate();

            if (result != null && result.Length > 0)
            {
                string contetntType = @"application/excel";
                HttpContext.Response.ContentType = contetntType;
                var content = new FileContentResult(result, contetntType);
                content.FileDownloadName = "file.xlsx";
                return content;
            }
            return NoContent();
        }
    }
}