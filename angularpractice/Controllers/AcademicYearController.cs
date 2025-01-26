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
        public async Task<IActionResult> Get()
        {
            await Task.Delay(2000);
            return Ok(academicYears);
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
