using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using CourseRadioPlan.Models;
using Microsoft.AspNetCore.Http;
using CourseRadioPlan.Services;

namespace CourseRadioPlan.Controllers
{
    public class HomeController : Controller
    {
        private readonly IRadioPlanSheetService _radioPlanSheetService;
        private readonly ILogger<HomeController> _logger;

        public HomeController(ILogger<HomeController> logger,
            IRadioPlanSheetService radioPlanSheetService)
        {
            _logger = logger;
            this._radioPlanSheetService = radioPlanSheetService;
        }

        public IActionResult Index([FromForm]HomeViewModel hvm = null)
        {
            if (hvm == null)
            {
                hvm = new HomeViewModel();
            }
            if (hvm.ExcelFile != null)
            {
                hvm.CourseName = this._radioPlanSheetService.GenerateFromFile(hvm.ExcelFile);
                return View(hvm);
            }
            return View(hvm);
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
