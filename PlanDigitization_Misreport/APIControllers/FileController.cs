using LessonLearntPortalWeb.Models;
using LessonLearntPortalWeb.Repository;
using LessonLearntPortalWeb.ViewModel;
using Microsoft.AspNetCore.Cors;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

// For more information on enabling Web API for empty projects, visit https://go.microsoft.com/fwlink/?LinkID=397860

namespace LessonLearntPortalWeb.APIControllers
{
    [EnableCors()]
    [Route("api/[controller]/[action]")]
    [ApiController]
    public class FileController : ControllerBase
    {

        private ExcelReportRepo repo;
        private IHostingEnvironment _hostingEnvironment;
        public FileController(ExcelReportRepo excelReportRepo, IHostingEnvironment hostingEnvironment)
        {
            repo = excelReportRepo;

            _hostingEnvironment = hostingEnvironment;
        }


        [HttpPost]
        public IActionResult DownloadReport(ExcelReportViewModel model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    if (repo.IsDatabaseOnline(model))
                    {
                        var excelBytes = repo.DownloadExcel(model);
                        if (excelBytes != null)
                        {
                            return File(excelBytes.FileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", excelBytes.FileName);
                        }
                        else
                        {
                            ApiResponce apiResponce = new ApiResponce();
                            apiResponce.Status = false;
                            apiResponce.Message = "Error downloading file";
                            return Ok(apiResponce);
                        }
                    }
                    else
                    {
                        ApiResponce apiResponce = new ApiResponce();
                        apiResponce.Status = false;
                        apiResponce.Message = "Database is not online";
                        return Ok(apiResponce);
                    }
                }
                else
                {
                    ApiResponce apiResponce = new ApiResponce();
                    apiResponce.Status = false;
                    apiResponce.Message = "Invalid Inputes";
                    return Ok(apiResponce);
                }
            }
            catch (Exception ex)
            {
                ApiResponce apiResponce = new ApiResponce();
                apiResponce.Status = false;
                apiResponce.Message = ex.Message;
                return Ok(apiResponce);
            }

        }

        [HttpPost]
        public IActionResult DownloadFilePath(ExcelReportViewModel model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    if (repo.IsDatabaseOnline(model))
                    {
                        var excelName = repo.DownloadExcelPath(model);
                        if (excelName != null)
                        {
                            ApiResponce apiResponce = new ApiResponce();
                            apiResponce.Status = true;
                            apiResponce.Message = "File Generated";
                            apiResponce.Data = excelName;
                            return Ok(apiResponce);
                        }
                        else
                        {
                            ApiResponce apiResponce = new ApiResponce();
                            apiResponce.Status = false;
                            apiResponce.Message = "Error Downloading File";
                            return Ok(apiResponce);
                        }
                    }
                    else
                    {
                        ApiResponce apiResponce = new ApiResponce();
                        apiResponce.Status = false;
                        apiResponce.Message = "Database Is Not Online";
                        return Ok(apiResponce);
                    }
                }
                else
                {
                    ApiResponce apiResponce = new ApiResponce();
                    apiResponce.Status = false;
                    apiResponce.Message = "Invalid Inputes";
                    return Ok(apiResponce);
                }
            }
            catch (Exception ex)
            {
                ApiResponce apiResponce = new ApiResponce();
                apiResponce.Status = false;
                apiResponce.Message = ex.Message;
                return Ok(apiResponce);
            }

        }

        [HttpPost]
        public IActionResult SendEmail(ExcelReportViewModel model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    if (repo.IsDatabaseOnline(model))
                    {
                        var response = repo.SendEmail(model);
                        return Ok(response);
                    }
                    else
                    {
                        return Ok(new EmailResponse() { Message = "Data base is not connected" });
                    }
                }
                else
                {
                    return Ok(new EmailResponse() { Message = "Please select all fields" });
                }
            }
            catch (Exception ex)
            {
                ApiResponce apiResponce = new ApiResponce();
                apiResponce.Status = false;
                apiResponce.Message = ex.Message;
                return Ok(apiResponce);
            }

        }


        [HttpPost]
        public IActionResult DownloadFileFromPath(ExcelReportViewModel model)
        {
            try
            {
                if (ModelState.IsValid)
                {
                    if (repo.IsDatabaseOnline(model))
                    {
                        ApiResponce apiResponce = new ApiResponce();
                        if (!String.IsNullOrEmpty(model.FileName))
                        {
                            var fileDownloadModel = repo.GetFileFromFilePath(model.FileName);
                            if (fileDownloadModel.FileBytes != null)
                            {
                                return File(fileDownloadModel.FileBytes, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", fileDownloadModel.FileName);
                            }
                            else
                            {
                                apiResponce.Status = false;
                                apiResponce.Message = "File Path is Invalid or File is not available";
                                return Ok(apiResponce);
                            }
                        }
                        else
                        {
                            apiResponce.Status = false;
                            apiResponce.Message = "File Path can not be Null or Empty";
                            return Ok(apiResponce);
                        }
                    }
                    else
                    {
                        return Ok(new EmailResponse() { Message = "Data base is not connected" });
                    }
                }
                else
                {
                    return Ok(new EmailResponse() { Message = "Please select all fields" });
                }
            }
            catch (Exception ex)
            {
                ApiResponce apiResponce = new ApiResponce();
                apiResponce.Status = false;
                apiResponce.Message = ex.Message;
                return Ok(apiResponce);
            }

        }

        [HttpPost]
        public IActionResult GetConnectionString(string SqlConnectionString)
        {

            ApiResponce apiResponce = new ApiResponce();

            try
            {

                apiResponce.Status = true;
                apiResponce.Message = "Success";
                apiResponce.Data = SqlConnectionString;
                return Ok(apiResponce);

            }
            catch (Exception ex)
            {
                apiResponce.Status = false;
                apiResponce.Message = ex.Message;
                return Ok(apiResponce);
            }
        }

    }
}
