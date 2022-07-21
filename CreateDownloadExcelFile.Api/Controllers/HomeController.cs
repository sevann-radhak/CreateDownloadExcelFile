using ClosedXML.Excel;
using CreateDownloadExcelFile.Api.Models;
using Microsoft.AspNetCore.Mvc;

namespace CreateDownloadExcelFile.Api.Controllers
{
    [ApiController]
    [Route("[controller]")]
    public class HomeController : ControllerBase
    {
        private readonly List<Student> _students = new();

        public HomeController()
        {
            for (int i = 0; i < 21; i++)
            {
                _students.Add(new()
                {
                    Id = i,
                    Name = $"Studend Name: {i}",
                    Roll = $"100{i}"
                });
            }
        }

        [HttpGet(Name = "CreateAndDownloadFile")]
        public IActionResult CreateAndDownloadFile()
        {
            using XLWorkbook? workbook = new();
            IXLWorksheet? worksheet = workbook.Worksheets.Add("Students");
            int currentRow = 1;

            #region Header
            worksheet.Cell(currentRow, 1).Value = "StudentId";
            worksheet.Cell(currentRow, 2).Value = "Name";
            worksheet.Cell(currentRow, 3).Value = "Roll";
            #endregion

            #region Body
            _ = _students
                .Select(s =>
                {
                    currentRow++;
                    worksheet.Cell(currentRow, 1).Value = s.Id;
                    worksheet.Cell(currentRow, 2).Value = s.Name;
                    worksheet.Cell(currentRow, 3).Value = s.Roll;

                    return s;
                })
                .ToList();
            #endregion

            using MemoryStream? stream = new();
            workbook.SaveAs(stream);
            byte[]? content = stream.ToArray();

            return File(
                content,
                "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                $"Students {DateTime.UtcNow}.xlsx");
        }
    }
}