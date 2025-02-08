using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace angularpractice.Controllers
{
    [Route("api/[controller]")]
    public class AcademicYearController : Controller
    {
        private readonly List<AcademicYear> academicYears;

        public AcademicYearController()
        {
            academicYears = new List<AcademicYear>
            {
                new AcademicYear { AcademicYearId = 1, Name = "2019", StartDate = new DateTime(2019, 1, 1), EndDate = new DateTime(2020, 12, 31) },
                new AcademicYear { AcademicYearId = 2, Name = "2020", StartDate = new DateTime(2020, 1, 1), EndDate = new DateTime(2021, 12, 31) },
                new AcademicYear { AcademicYearId = 3, Name = "2021", StartDate = new DateTime(2021, 1, 1), EndDate = new DateTime(2022, 12, 31) }
            };
        }

        [HttpGet]
        public async Task<IActionResult> Get(int id, CancellationToken cancellationToken)
        {
            try
            {
                // Simulamos una tarea de larga duración
                for (int i = 0; i < 10; i++)
                {
                    // Simulamos trabajo
                    await Task.Delay(1000, cancellationToken);

                    // Podemos agregar logging para ver el progreso
                    System.Diagnostics.Debug.WriteLine($"Iteración {i + 1} completada");
                }
                return Ok(academicYears);
            }
            catch (OperationCanceledException)
            {
                // Logging de la cancelación
                System.Diagnostics.Debug.WriteLine("La operación fue cancelada por el cliente");
                return StatusCode(499, "Cliente cerró la conexión");
            }
        }

        [HttpGet("generar")]
        public IActionResult GenerarExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Hoja1");

                // Agregar datos al Excel
                worksheet.Cell("A1").Value = "Nombre";
                worksheet.Cell("B1").Value = "Edad";
                worksheet.Cell("A2").Value = "Juan";
                worksheet.Cell("B2").Value = 30;
                worksheet.Cell("A3").Value = "María";
                worksheet.Cell("B3").Value = 25;

                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Seek(0, SeekOrigin.Begin);

                    // Agregar cabeceras para la descarga
                    Response.Headers.Add("Content-Disposition", "attachment; filename=\"mi_excel.xlsx\"");
                    Response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                    return new FileContentResult(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                }
            }
        }
    }

    public class AcademicYear
    {
        public int AcademicYearId { get; set; }
        public string Name { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime EndDate { get; set; }
    }
}
