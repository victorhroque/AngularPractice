using ClosedXML.Excel;
using Microsoft.AspNetCore.Mvc;

namespace angularpractice.Controllers
{
    public class ExcelController : Controller
    {
        [HttpGet("generar")]
        public IActionResult GenerarExcel()
        {
            using (var workbook = new XLWorkbook())
            {
                var worksheet = workbook.Worksheets.Add("Informe de Ventas");

                // Datos de ejemplo
                var ventas = new[]
                {
                new { Producto = "Camiseta", Cantidad = 10, Precio = 20 },
                new { Producto = "Pantalón", Cantidad = 5, Precio = 30 },
                new { Producto = "Zapatos", Cantidad = 3, Precio = 50 },
            };

                // Encabezados
                worksheet.Cell("A1").Value = "Producto";
                worksheet.Cell("B1").Value = "Cantidad";
                worksheet.Cell("C1").Value = "Precio";
                worksheet.Cell("D1").Value = "Total";

                // Estilos para encabezados
                worksheet.Range("A1:D1").Style.Font.Bold = true;
                worksheet.Range("A1:D1").Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                worksheet.Range("A1:D1").Style.Fill.BackgroundColor = XLColor.LightGray;

                // Datos
                for (int i = 0; i < ventas.Length; i++)
                {
                    worksheet.Cell(i + 2, 1).Value = ventas[i].Producto;
                    worksheet.Cell(i + 2, 2).Value = ventas[i].Cantidad;
                    worksheet.Cell(i + 2, 3).Value = ventas[i].Precio;
                    worksheet.Cell(i + 2, 4).FormulaA1 = $"=B{i + 2}*C{i + 2}";
                }

                // Formato de moneda
                worksheet.Column(3).Style.NumberFormat.Format = "#,##0.00";
                worksheet.Column(4).Style.NumberFormat.Format = "#,##0.00";

                // Suma total
                var ultimaFila = ventas.Length + 2;
                worksheet.Cell(ultimaFila, 4).FormulaA1 = $"=SUM(D2:D{ultimaFila - 1})";
                worksheet.Cell(ultimaFila, 4).Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Right;
                worksheet.Cell(ultimaFila, 4).Style.Font.Bold = true;


                // Guardar y descargar
                using (var stream = new MemoryStream())
                {
                    workbook.SaveAs(stream);
                    stream.Seek(0, SeekOrigin.Begin);

                    Response.Headers.Add("Content-Disposition", "attachment; filename=\"informe_ventas.xlsx\"");
                    Response.Headers.Add("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

                    return new FileContentResult(stream.ToArray(), "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                }
            }
        }
    }
}
