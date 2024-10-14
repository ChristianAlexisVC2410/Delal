using BussinessLogic;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Net.Mime;
using System.Reflection.Metadata.Ecma335;

namespace GestorDelal.Controllers
{
    [Route("api/Delal")]
    [ApiController]
    public class DelalController : ControllerBase
    {
        private readonly DelalLogic _delalLogic;

        public DelalController(DelalLogic delalLogic)
        {
            _delalLogic = delalLogic;
        }

        [HttpPost("cargaArchivo")]
        public async Task<ActionResult<object>> CrearArchivoExcel()
        {
            try
            {
                var fi = Request;
                IFormFile files = Request.Form.Files.GetFile("file");
                (object listas,MemoryStream archivoExcel) = await _delalLogic.CargarArchivoLogic(files);
                var fileName = "GeneratedExcel.xlsx";
                var contentType = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
                return File(archivoExcel, contentType, fileName);
            }
            catch (Exception ex)
            {
                return BadRequest(ex.Message);
            }
        }
    }
}
