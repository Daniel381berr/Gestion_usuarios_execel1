using GESTIÓN_DE_USUARIOS.Models;
using Microsoft.AspNetCore.Mvc;
using System.Diagnostics;

using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using NPOI.HSSF.UserModel;
using GESTIÓN_DE_USUARIOS.Models.VIEW_MODELS;
using System.Diagnostics.Eventing.Reader;

namespace GESTIÓN_DE_USUARIOS.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger)
        {
            _logger = logger;
        }

        public IActionResult Index()
        {
            return View();
        }

        [HttpPost]
        public IActionResult MostrarDatos([FromForm]IFormFile ArchivoExel)
        {
            Stream stream = ArchivoExel.OpenReadStream();
            IWorkbook MiExcel = null;

            if (Path.GetExtension(ArchivoExel.FileName) == ".xlsx")
            {
                MiExcel = new XSSFWorkbook(stream);
            }
            else {

                MiExcel = new HSSFWorkbook(stream);

            }
            
            ISheet HojaExcel = MiExcel.GetSheetAt(0);

            int cantidadFilas = HojaExcel.LastRowNum;

            List<VMContactos> lista = new List<VMContactos>();

            for (int i = 1; i <= cantidadFilas; i++)
            {

                IRow fila = HojaExcel.GetRow(i);

                lista.Add(new VMContactos
                {
                    NOMBRE = fila.GetCell(0).ToString(),
                    APELLIDO = fila.GetCell(1).ToString(),
                    EDAD = fila.GetCell(2).ToString(),
                    PUESTO = fila.GetCell(3).ToString(),

                });


            }

            return StatusCode(StatusCodes.Status200OK,lista);
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }
    }
}