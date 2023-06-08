using LessonLearntPortalWeb.Repository;
using LessonLearntPortalWeb.ViewModel;
using Microsoft.AspNetCore.Http;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using System.Collections.Generic;

namespace LessonLearntPortalWeb.Controllers
{
    public class DefaultController : Controller
    {
        private ExcelReportRepo repo;
        public DefaultController(ExcelReportRepo excelReportRepo)
        {
            repo = excelReportRepo;
        }

        public IActionResult IndexData()
        {
            return View();
        }

        // GET: Default
        public IActionResult Index()
        {
            ExcelReportViewModel model = new ExcelReportViewModel();
            
            model.Date = System.DateTime.Today.AddDays(-1);
            model.SqlConnectionString = "";
            return View(model);
        }

        //[WebMethod]
        [HttpPost]
        public IActionResult DownloadReport(ExcelReportViewModel model)
        {
            if (ModelState.IsValid)
            {
               // ExcelReportRepo repo = new ExcelReportRepo();
                if (repo.IsDatabaseOnline(model))
                {
                    var excelBytes = repo.DownloadExcel(model);
                    if (excelBytes != null)
                    {                     
                        return File(excelBytes.FileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelBytes.FileName);                      
                    }
                    else
                    {
                        return Json(new DownloadReportResponse { Message = "Error downloading file" });
                    }
                }
                else
                {
                    return Json(new DownloadReportResponse { Message = "Database is not online" });
                }
            }
            else
            {
                return Json(new DownloadReportResponse { Message = "Invalid inputs" });
            }
        }

        [HttpPost]
        public IActionResult SendEmail(ExcelReportViewModel model)
        {
            if (ModelState.IsValid)
            {
                //ExcelReportRepo repo = new ExcelReportRepo();
                if (repo.IsDatabaseOnline(model))
                {
                    var response = repo.SendEmail(model);
                    return Json(response);
                }
                else
                {
                    return Json(new EmailResponse() { Message = "Data base is not connected" });
                }
            }
            else
            {
                return Json(new EmailResponse() { Message = "Please select all fields" });
            }
        }

    }
}
