using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.Threading.Tasks;
using VormerProject.Services;

namespace VormerProject.Controllers
{
    [ApiController]
    [Route("api/[controller]")]
    public class ExcelUploadController : ControllerBase
    {
        private readonly IExcelService _excelService;

        public ExcelUploadController(IExcelService excelService)
        {
            _excelService = excelService;
        }

        [HttpPost("upload")]
        public async Task<IActionResult> UploadExcel(IFormFile file)
        {
            if (file == null || file.Length == 0)
                return BadRequest("File verplicht.");

            var jsonResult = await _excelService.ReturnJson(file);
            return Ok(jsonResult);
        }
    }
}

