using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Mvc.Rendering;
using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;

namespace LessonLearntPortalWeb.ViewModel
{
    public class ExcelReportViewModel
    {
        [Required]
        [DisplayFormat(DataFormatString = "{0:yyyy-mm-dd}", ApplyFormatInEditMode = true)]
        public DateTime Date { get; set; }
        [Required]
        public string CompanyCode { get; set; } //= "TEAL_DTVS";
        public string CompanyName { get; set; }
        public List<SelectListItem> CompanyCodeList { get; set; }
        [Required]
        public string PlantCode { get; set; } //= "TEAL_DTVS01";
        public List<SelectListItem> PlantCodeList { get; set; }
        [Required]
        public string LineCode { get; set; } //= "VCTM01";
        public List<SelectListItem> LineCodeList { get; set; }
        public string StationCode { get; set; }  //= "M1"
        public List<SelectListItem> StationCodeList { get; set; }

        public string SqlConnectionString { get; set; } //= "Connection String"

        public string FileName { get; set; }
    }

    public class EmailResponse
    {
        public bool IsSent { get; set; }
        public string Message { get; set; }
        public string[] ToSent { get; set; }
        public string[] CCSent { get; set; }
        public string[] BCCSent { get; set; }
    }
    public class DownloadReportResponse
    {
        public string Message { get; set; }
    }

    public class FileDownloadModel
    {
        public byte[] FileBytes { get; set; }
        public string FileName { get; set; }
        public string fileFullPathName { get; set; } 
    }

    public class DownloadFileNameModel 
    {
        public FileContentResult File { get; set; }
        public string FileName { get; set; }
    }
    public class ApiResponce
    {
        public bool Status { get; set; }
        public string Message { get; set; }
        public object Data { get; set; }
    }
    //public class FilePathModel
    //{
    //    public string FileName { get; set; }
    //}

}