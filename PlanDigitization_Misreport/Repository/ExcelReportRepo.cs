//using Antlr.Runtime;
//using ClosedXML.Excel;
//using DocumentFormat.OpenXml.EMMA;
//using DocumentFormat.OpenXml.Office2010.ExcelAc;
//using ClosedXML.Excel;
using ClosedXML.Excel;
using LessonLearntPortalWeb.ViewModel;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Mvc.Rendering;
using Microsoft.Extensions.Configuration;
using System;
using System.Collections.Generic;
using System.Configuration;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Net.Mail;
using LessonLearntPortalWeb.Models;
//using System.Web.Hosting;

namespace LessonLearntPortalWeb.Repository
{

    //this method is receiving the sql string from the front end, getting the excel file, sql connection, getting the data from database and attaching the mail body
    public class ExcelReportRepo
    {
        private string _connection;
   
        #region Property  
        private IHostingEnvironment _hostingEnvironment;
        private readonly IConfiguration _configuration;
        #endregion

        public List<SelectListItem> CompanyCodeList;
        public List<SelectListItem> PlantCodeList;
        public List<SelectListItem> LineCodeList;
        public List<SelectListItem> StationCodeList;
        private string ExcelFilePath;
        private string Historical_ExcelFiles;
        private string TemplatePath;

        public ExcelReportRepo(IHostingEnvironment hostingEnvironment, IConfiguration configuration)
        {
            _hostingEnvironment = hostingEnvironment;
            var tt = Directory.GetCurrentDirectory();
            var ttrtrt = this._hostingEnvironment.WebRootPath.ToString();
            ExcelFilePath = Path.Combine(_hostingEnvironment.WebRootPath, "UploadFiles", "ExcelFiles");
            Historical_ExcelFiles = Path.Combine(this._hostingEnvironment.WebRootPath, "UploadFiles", "Historical_ExcelFiles");
            TemplatePath = Path.Combine(this._hostingEnvironment.WebRootPath, "UploadFiles", "Template");

        }

        private string GetConnectionString()
        {
            string constr = "";
            constr = _connection;
            return constr;
        }

        // This method is verify the database connection is connected or not 
        public bool IsDatabaseOnline(ExcelReportViewModel model)
        {
            bool isConnected = false;
            _connection = model.SqlConnectionString;
            using (SqlConnection connection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    connection.Open();
                    isConnected = true;
                }
                catch (Exception ex)
                {
                    isConnected = false;
                }
                finally
                {
                    if (connection != null)
                    {
                        connection.Close();
                      
                    }
                }
            }
            if (isConnected == true )
            {
               // GetMasterCodes();
            }
            return isConnected;
        }

        // This method is get the master details (Company code,plant code,Line code & station code) from the received connecting string value 

        //public string GetMasterCodes()
        //{

        //    var status = "";
        //    DataSet ds = new DataSet();

        //    using (SqlConnection con = new SqlConnection(GetConnectionString()))
        //    {
        //        try
        //        {

        //            con.Open();
        //            SqlCommand cmd = new SqlCommand("SP_MasterCodes_Get", con);
        //            cmd.CommandType = CommandType.StoredProcedure;
        //            cmd.Parameters.AddWithValue("@strCodeType", "All");
        //            SqlDataAdapter holidayDa = new SqlDataAdapter(cmd);
        //            cmd.CommandTimeout = 0;
        //            SqlDataAdapter da = new SqlDataAdapter(cmd);
        //            da.Fill(ds);
        //            if (ds.Tables.Count > 0)
        //            {

        //                if (ds.Tables[0].Rows.Count > 0)
        //                {
        //                    CompanyCodeList = new List<SelectListItem>();
        //                    SelectListItem obj = new SelectListItem();
        //                    foreach (DataRow dr in ds.Tables[0].Rows)
        //                    {
        //                        obj = new SelectListItem();
        //                        obj.Value = dr.ItemArray[0].ToString();
        //                        CompanyCodeList.Add(obj);
        //                    }
        //                }

        //                if (ds.Tables[1].Rows.Count > 0)
        //                {
        //                    PlantCodeList = new List<SelectListItem>();
        //                    SelectListItem obj = new SelectListItem();
        //                    foreach (DataRow dr in ds.Tables[1].Rows)
        //                    {
        //                        obj = new SelectListItem();
        //                        obj.Value = dr.ItemArray[0].ToString();
        //                        PlantCodeList.Add(obj);
        //                    }
        //                }

        //                if (ds.Tables[2].Rows.Count > 0)
        //                {
        //                    LineCodeList = new List<SelectListItem>();
        //                    SelectListItem obj = new SelectListItem();
        //                    foreach (DataRow dr in ds.Tables[2].Rows)
        //                    {
        //                        obj = new SelectListItem();
        //                        obj.Value = dr.ItemArray[0].ToString();
        //                        LineCodeList.Add(obj);
        //                    }
        //                }

        //                if (ds.Tables[3].Rows.Count > 0)
        //                {
        //                    StationCodeList = new List<SelectListItem>();
        //                    SelectListItem obj = new SelectListItem();
        //                    foreach (DataRow dr in ds.Tables[3].Rows)
        //                    {
        //                        obj = new SelectListItem();
        //                        obj.Value = dr.ItemArray[0].ToString();
        //                        StationCodeList.Add(obj);
        //                    }
        //                }
        //            }
        //            return status;
        //        }
        //        catch (Exception ex)
        //        {
        //            throw ex;
        //        }
        //        finally
        //        {
        //            con.Close();
        //        }
        //    }

        //}

        //This method is used for the if excel file already available from directory then get the file name not available then Generate Excel Report
        public FileDownloadModel DownloadExcel(ExcelReportViewModel model)
        {
            try {
                string MonthLabel = model.Date.ToString("MMMM_yyyy");
                string date = model.Date.ToString("dd");
                string day = model.Date.DayOfWeek.ToString();
                string lineCode = model.LineCode;
                string stationCode = model.StationCode;
                string fileName = MonthLabel + "_Day_" + date + "_" + lineCode + "_" + stationCode + "_Report.xlsx";
                string filePathName = "";
                var fileBytes = GetFileIfExist(fileName, lineCode, stationCode, out filePathName);
                var fileDownloadModel = new FileDownloadModel();

                if (fileBytes != null)
                {
                    fileDownloadModel.FileBytes = fileBytes;
                    fileDownloadModel.FileName = fileName;
                    fileDownloadModel.fileFullPathName = filePathName;
                    return fileDownloadModel;
                }
                else
                {
                    var fname = "";
                    fileDownloadModel.FileBytes = GenerateExcelReport(model, out fname);
                    fileDownloadModel.fileFullPathName = fname;
                    fileDownloadModel.FileName = fileName;
                    fileDownloadModel.fileFullPathName = filePathName;
                    return fileDownloadModel;
                }
            }
            catch(Exception ex)
            {
                throw ex;
            }           
        }

        //This method is used  to checking the Historical file if selected date file is available or not 
        private byte[] GetFileIfExist(string name, string lineCode, string stationCode, out string filePathName)
        {
            string path = Path.Combine(_hostingEnvironment.WebRootPath, "UploadFiles", "Historical_ExcelFiles", lineCode, stationCode, name);
            if (File.Exists(path))
            {
                filePathName = path;
                return File.ReadAllBytes(path);
            }
            filePathName = string.Empty;
            return null;
        }

        //This method is used to Generate Excel Report based on download button
        private byte[] GenerateExcelReport(ExcelReportViewModel model, out string fileFullPathName)
        {
            DateTime today = model.Date;
            var dat = today.ToString("yyyy-MM-dd");
            DateTime nxttoday = today.AddDays(1);
            var nxtdat = nxttoday.ToString("yyyy-MM-dd");
            string holidayname = "";
            string companyCode = model.CompanyCode;
            string companyName = model.CompanyName;
            string plantCode = model.PlantCode;
            string lineCode = model.LineCode;
            string stationCode = model.StationCode;


            using (SqlConnection connection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    SqlCommand holiDayCmd = new SqlCommand("SELECT [HolidayReason],[Date] FROM [dbo].[tbl_Holiday] where Date= @date and CompanyCode= @CompanyCode and PlantID= @PlantCode ", connection);
                    holiDayCmd.Parameters.AddWithValue("@date", dat);
                    holiDayCmd.Parameters.AddWithValue("@CompanyCode", companyCode);
                    holiDayCmd.Parameters.AddWithValue("@PlantCode", plantCode);
                    SqlDataAdapter holidayDa = new SqlDataAdapter(holiDayCmd);
                    DataTable holidayDt = new DataTable();
                    holidayDa.Fill(holidayDt);
                    connection.Open();
                    var first = new DateTime(today.Year, today.Month, 1);
                    var frstdaystr = first.ToString("yyyy-MM-dd");
                    if (dat == frstdaystr)
                    {
                        SqlCommand cmd555 = new SqlCommand("truncate table [dbo].[DailyReport_Time];", connection);
                        cmd555.ExecuteNonQuery();
                    }
                    else
                    {
                        SqlCommand cmd555 = new SqlCommand("delete from [dbo].[DailyReport_Time] where Date=@date;", connection);
                        cmd555.Parameters.AddWithValue("@date", dat);
                        cmd555.ExecuteNonQuery();
                    }

                    if (holidayDt.Rows.Count != 0)
                    {
                        holidayname = holidayDt.Rows[0][0].ToString();
                        fileFullPathName = "";
                        return null;
                    }
                    else
                    {
                        SqlCommand assetCmd = new SqlCommand("select AssetID as Machine_Code, AssetName as MachineName,f.FunctionID as Line_code" +
                        //" from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID\", connection);
                        " from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID where f.FunctionID=@FunctionID AND AssetID=@AssetID", connection);
                        assetCmd.Parameters.AddWithValue("@FunctionID", lineCode);
                        assetCmd.Parameters.AddWithValue("@AssetID", stationCode);
                        SqlDataAdapter assetsDa = new SqlDataAdapter(assetCmd);
                        DataTable assetsDt = new DataTable();
                        assetsDa.Fill(assetsDt);
                        string[] name = new string[assetsDt.Rows.Count];
                        string[] line = new string[assetsDt.Rows.Count];
                        string[] machinecode = new string[assetsDt.Rows.Count];
                        for (int i = 0; i < assetsDt.Rows.Count; i++)
                        {
                            machinecode[i] = assetsDt.Rows[i][0].ToString();
                            name[i] = assetsDt.Rows[i][1].ToString();
                            line[i] = assetsDt.Rows[i][2].ToString();
                        }
                        string day = model.Date.DayOfWeek.ToString();
                        string templatePath = Path.Combine(this._hostingEnvironment.WebRootPath, "UploadFiles", "Template", "Template_3.xlsx");
                        String MonthLabel = today.ToString("MMMM_yyyy");
                        String dayy = today.ToString("dd");
                        var excelFilePathWithLineCode = Path.Combine(ExcelFilePath, lineCode, stationCode);
                        if (!Directory.Exists(excelFilePathWithLineCode))
                        {
                            Directory.CreateDirectory(excelFilePathWithLineCode);
                        }
                        var datefile = model.Date.ToString("dd");
                        string fileName = MonthLabel + "_Day_" + datefile + "_" + lineCode + "_" + stationCode + "_Report.xlsx";
                        string filepath = Path.Combine(excelFilePathWithLineCode, fileName);
                        fileFullPathName = filepath;
                        XLWorkbook workbook = new XLWorkbook(templatePath);

                        workbook.SaveAs(filepath);
                        for (int i = 0; i < assetsDt.Rows.Count; i++)
                        {
                            DataSet dataset = getDataSet(GetConnectionString(), line[i], today, companyCode, plantCode, machinecode[i]);
                            if (i == 0)
                            {
                                LoopAllMachines GF1 = new LoopAllMachines();
                                //Class1 GF = new Class1();
                                GF1.getData(GetConnectionString(), dataset, name[i], templatePath, filepath, line[i], dat);
                            }
                            GetExcelFile GF = new GetExcelFile();
                            GF.getData(GetConnectionString(), dataset, name[i], templatePath, filepath, line[i], companyCode, companyName, plantCode, machinecode[i], dat, i);
                        }


                        //DataSet dataset = getDataSet(GetConnectionString(), lineCode, today, companyCode, plantCode, stationCode);
                        ////if (i == 0)
                        ////{
                        //    LoopAllMachines GF1 = new LoopAllMachines();
                        //    //Class1 GF = new Class1();
                        //    GF1.getData(GetConnectionString(), dataset, name[0], templatePath, filepath, lineCode, dat);
                        ////}
                        //GetExcelFile GF = new GetExcelFile();
                        //GF.getData(GetConnectionString(), dataset, name[0], templatePath, filepath, lineCode, companyCode, plantCode, stationCode, dat, 0);



                        byte[] fileContent = File.ReadAllBytes(filepath);
                        DirectoryInfo dirInfo = new DirectoryInfo(excelFilePathWithLineCode);
                        foreach (FileInfo excelFile in dirInfo.EnumerateFiles())
                        {
                            var historicalExcelFilePathWithLineCode = Path.Combine(Historical_ExcelFiles, lineCode, stationCode);
                            if (!Directory.Exists(historicalExcelFilePathWithLineCode))
                            {
                                Directory.CreateDirectory(historicalExcelFilePathWithLineCode);
                            }

                            if (!File.Exists(Path.Combine(historicalExcelFilePathWithLineCode, MonthLabel + "_Day_" + dayy + "_" + lineCode + "_" + stationCode + "_Report.xlsx")))
                            {
                                File.Copy(filepath, Path.Combine(historicalExcelFilePathWithLineCode, MonthLabel + "_Day_" + dayy + "_" + lineCode + "_" + stationCode + "_Report.xlsx"));  
                            }
                        }
                        return fileContent;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    connection.Close();
                }
            }
        }

        //This method is used to download excel file and get dataset from the sql database
        public DataSet getDataSet(string connStr, string line, DateTime today, string companyCode, string plantCode, string machinecode)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("sp_Project_T_MIS_Report", con);
                    cmd.CommandType = CommandType.StoredProcedure;
                    cmd.CommandTimeout = 0;
                    cmd.Parameters.Add("@Machine_Code", SqlDbType.NVarChar, 150).Value = machinecode;
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.NVarChar, 150).Value = companyCode;
                    cmd.Parameters.Add("@Line_code", SqlDbType.NVarChar, 150).Value = line;                    
                    cmd.Parameters.Add("@Date", SqlDbType.NVarChar, 150).Value = today.ToString("yyyy-MM-dd");   
                    cmd.Parameters.Add("@PlantCode", SqlDbType.NVarChar, 150).Value = plantCode;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    return (ds);
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally

                {
                    con.Close();
                    ds.Dispose();
                }
            }
        }



        //This method sending the mail with excel report attachements to 'bcc' and 'cc' mails 
        public EmailResponse SendEmail(ExcelReportViewModel model)
        {
            string[] strEmailRespose = new string[2];
            try
            {
                string MonthLabel = model.Date.ToString("MMMM_yyyy");
                string date = model.Date.ToString("dd");
                string day = model.Date.DayOfWeek.ToString();
                string companyCode = model.CompanyCode;
                string plantCode = model.PlantCode;
                string lineCode = model.LineCode;
                string stationCode = model.StationCode;
                string fileName = MonthLabel + "_Day_" + date + "_" + lineCode + "_" + stationCode + "_Report.xlsx";
                var searchFile = Path.Combine(Historical_ExcelFiles, lineCode, stationCode);
                byte[] attachmentFile = null;
                bool isFileExists = Directory.Exists(searchFile);

                if (isFileExists == false)
                {
                    Directory.CreateDirectory(searchFile);
                }

                string[] files = Directory.GetFiles(searchFile, fileName);
                if (files.Count() != 0)
                {
                    foreach (var file in files)
                    {
                        byte[] currentFile = File.ReadAllBytes(file);
                        attachmentFile = currentFile;
                    }
                }
                else
                {
                    var details = DownloadExcel(model);
                    attachmentFile = details.FileBytes;
                    fileName = details.FileName;
                }

                byte[] fileBytes = attachmentFile != null ? attachmentFile : null;

                if (attachmentFile == null)
                {
                    return new EmailResponse() { IsSent = false, Message = "Please,There is no attachment" };
                }

                string[] toEmailAry = getemailDataSet(companyCode, plantCode, lineCode, stationCode);
                string[] ccEmailAry = getccDataSet(companyCode, plantCode, lineCode, stationCode);
                string[] bccEmailAry = getbccDataSet(companyCode, plantCode, lineCode, stationCode);
                if (toEmailAry.Length <= 0)
                {
                    return new EmailResponse() { IsSent = false, Message = "To email ids are not configured" };
                }
                if (ccEmailAry.Length <= 0)
                {
                    return new EmailResponse() { IsSent = false, Message = "cc email ids are not configured" };
                }
                string htmlBody = string.Empty;
                string textBody = "";
                string strMachineName = "";
                if (fileBytes == null)
                {
                    htmlBody = "<p>Production count of <b>" + plantCode + " - " + lineCode + ":" + model.Date.ToString("yyyy-MM-dd") + "</b></p>" +
                                    "<p><b>Day Start Date:" + model.Date.ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>" +
                                    "<p><b>Day End Date:" + model.Date.AddDays(1).ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>";
                }
                else
                {
                    htmlBody = "<p>Production count of <b>" + plantCode + " - " + lineCode + ":" + model.Date.ToString("yyyy-MM-dd") + "</b></p>" +
                                    "<p><b>Day Start Date:" + model.Date.ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>" +
                                    "<p><b>Day End Date:" + model.Date.AddDays(1).ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>";


                    textBody = " <table bordercolor='#00A8FC' border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 900 + "><tr  bordercolor='#00A8FC' bgcolor='#00A8FC'>" +
                        "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#FFFFFF'><b>Machine Name</b></td> " +
                        "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#FFFFFF'><b>Variant Name</b></td>" +
                        "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#FFFFFF'><b>Actual Ok Parts</b></td>" +
                        "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#FFFFFF'><b>Actual NOk Parts</b></td>" +
                        "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#FFFFFF'><b>Rejection %</b></td>" +
                        "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#FFFFFF'><b>UpTime (In Mins/in %)</b></td>" +
                        "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#FFFFFF'><b>Downtime (In Mins/in %)</b></td>" +
                        "</tr>";

                    string[] name;
                    string[] line;
                    string[] machinecode;
                    using (SqlConnection connection = new SqlConnection(GetConnectionString()))
                    {
                        try
                        {

                            SqlCommand assetCmd = new SqlCommand("select AssetID as Machine_Code, AssetName as MachineName,f.FunctionID as Line_code" +
                            " from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID where f.FunctionID=@FunctionID AND AssetID=@AssetID", connection);
                            assetCmd.Parameters.AddWithValue("@FunctionID", lineCode);
                            assetCmd.Parameters.AddWithValue("@AssetID", stationCode);
                            SqlDataAdapter assetsDa = new SqlDataAdapter(assetCmd);
                            DataTable assetsDt = new DataTable();
                            connection.Open();
                            assetsDa.Fill(assetsDt);
                            name = new string[assetsDt.Rows.Count];
                            line = new string[assetsDt.Rows.Count];
                            machinecode = new string[assetsDt.Rows.Count];
                            for (int i = 0; i < assetsDt.Rows.Count; i++)
                            {
                                machinecode[i] = assetsDt.Rows[i][0].ToString();
                                name[i] = assetsDt.Rows[i][1].ToString();
                                line[i] = assetsDt.Rows[i][2].ToString();
                            }
                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally

                        {
                            connection.Close();
                        }
                    }
                    

                    DataSet dataset = getDataSet(GetConnectionString(), lineCode, model.Date, companyCode, plantCode, stationCode);
                    
                    if (dataset.Tables[1].Rows.Count > 0 )
                    {

                        float strRejection = 0;
                       
                        for (int i = 0; i < name.Length; i++)
                        {
                            if (stationCode == machinecode[i] && lineCode == line[i])
                            {
                                strMachineName = name[i].ToString();
                            }
                        }

                        foreach (DataRow Row in dataset.Tables[1].Rows)
                        {
                            
                                if ((Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"])) != 0)
                                {
                                    strRejection = ((float)Math.Round((((float)(Convert.ToInt32(Row["actual_nok"])) / (float)(Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"]))) * 100) * 100f) / 100f);
                                }
                                else
                                {
                                    strRejection = 0;
                                }


                            
                                textBody += "<tr>" +
                                    "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'>" + strMachineName + "</td>" +
                                    "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'> " + Row["VariantCode"].ToString() + "</td> " +
                                    "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> " + Row["ActualProduction"].ToString() + "</td> " +
                                    "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> " + Row["actual_nok"].ToString() + "</td> " +
                                    "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#1A0303'> " + strRejection.ToString() + "</td> " +
                                    "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> " + Row["UpTime(min)"].ToString() + " Mins / "  + Row["UPTime%"].ToString()  + " %"   + " </td> " +
                                    "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> " + Row["DownTime(Min)"].ToString() + " Mins / " + Row["DownTime%"].ToString() + " %" + " </td> " +
                                    "</tr>";
                        }
                        textBody += "</table>";
                    }
                    else
                    {
                        for ( int i = 0; i < name.Length; i++)
                        {

                            string dayname="";
                            dayname = model.Date.ToString("dddd").ToString();
                            if (dayname == "Sunday")
                            {
                                textBody += "<tr>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'>" + name[i].ToString() + "</td>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'> --Sunday--Holiday-- </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "</tr>";
                            }
                            else
                            {
                                textBody += "<tr>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'>" + name[i].ToString() + "</td>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'> --Data Not Logged-- </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "</tr>";
                            }

                           
                        }
                        textBody += "</table>";
                       
                    }
                    textBody += "<p>For more dteails refer the attachment!</p>" +
                                    "<p>Refer the portal for more info <u style='color:#00A8FC'><a href='https://i4metrics.titan.in/'>" + " click  to login " + "</a></u></p>" +
                    "<p>***Mail generated from TEAL IIOT Portal Email App services***</p>";

                    htmlBody += textBody;
                }

                strEmailRespose = SendEmail(htmlBody, toEmailAry, ccEmailAry, bccEmailAry, model.Date, fileBytes, fileName, strMachineName);
                return new EmailResponse() { IsSent = Convert.ToBoolean(Convert.ToInt32(strEmailRespose[0])), Message = strEmailRespose[1], ToSent = toEmailAry, CCSent = ccEmailAry, BCCSent = bccEmailAry };
                   
            }
            catch (Exception ex)
            {
                if (strEmailRespose[0] == null )
                {
                    return new EmailResponse() { IsSent = false, Message = ex.ToString() };
                }
                else
                {
                    return new EmailResponse() { IsSent = Convert.ToBoolean(Convert.ToInt32(strEmailRespose[0])), Message = strEmailRespose[1].ToString() };
                }
                
            }
        }

        // this method is getting the TO mail ids from database
        public string[] getemailDataSet(string companyCode, string plantCode, string lineCode, string stationCode)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select * from tbl_emails " +
                        "where Companycode=@CompanyCode and Plantcode=@PlantCode and Status=@to and line_code=@linecode", con);
                    cmd.Parameters.Add("@to", SqlDbType.Int).Value = 1;
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = companyCode;
                    cmd.Parameters.Add("@PlantCode", SqlDbType.VarChar).Value = plantCode;
                    cmd.Parameters.Add("@linecode", SqlDbType.VarChar).Value = lineCode;
                    cmd.Parameters.Add("@stationCode", SqlDbType.VarChar).Value = stationCode;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    string[] emails = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow.Field<string>("Email_ID")).ToArray();
                    return emails;
                }
                catch(Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    ds.Dispose();
                    con.Close();
                }
            }
        }

        // this method is getting the CC mail ids from database
        public string[] getccDataSet(string companyCode, string plantCode, string lineCode, string stationCode)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select * from tbl_emails" +
                        " where Companycode=@CompanyCode and Plantcode=@PlantCode and Status=@cc and line_code=@linecode", con);
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = companyCode;
                    cmd.Parameters.Add("@PlantCode", SqlDbType.VarChar).Value = plantCode;
                    cmd.Parameters.Add("@cc", SqlDbType.Int).Value = 0;
                    cmd.Parameters.Add("@linecode", SqlDbType.VarChar).Value = lineCode;
                    cmd.Parameters.Add("@stationCode", SqlDbType.VarChar).Value = stationCode;

                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    string[] emails = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow.Field<string>("Email_ID")).ToArray();
                    return emails;
                }
                catch(Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    ds.Dispose();
                    con.Close();
                }
            }
        }

        // this method is getting the bcc mailIDs from the database
        public string[] getbccDataSet(string companyCode, string plantCode, string lineCode, string stationCode)
        {
            DataSet ds = new DataSet();
            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    con.Open();
                    SqlCommand cmd = new SqlCommand("select * from tbl_emails" +
                        " where Companycode=@CompanyCode and Plantcode=@PlantCode and Status=@bcc and line_code=@linecode", con);
                    cmd.Parameters.Add("@CompanyCode", SqlDbType.VarChar).Value = companyCode;
                    cmd.Parameters.Add("@PlantCode", SqlDbType.VarChar).Value = plantCode;
                    cmd.Parameters.Add("@bcc", SqlDbType.Int).Value = 2;
                    cmd.Parameters.Add("@linecode", SqlDbType.VarChar).Value = lineCode;
                    cmd.Parameters.Add("@stationCode", SqlDbType.VarChar).Value = stationCode;
                    SqlDataAdapter da = new SqlDataAdapter(cmd);
                    da.Fill(ds);
                    string[] emails = ds.Tables[0].AsEnumerable().Select(dataRow => dataRow.Field<string>("Email_ID")).ToArray();
                    return emails;
                }
                catch(Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    ds.Dispose();
                    con.Close();
                }
            }
        }

        //this method is get gmail settings,add To,CC,BCC Mail id and sending the mail with attachement and mail body

        public string[] SendEmail(string htmlBody, string[] MailToset, string[] CCset, string[] BCCset, DateTime date, byte[] attachmentFile, string fileName,string strMachineName)
        {

            using (SqlConnection con = new SqlConnection(GetConnectionString()))
            {
                string[] strArray = new string[2];

                try
                {
                    MailMessage mail = new MailMessage();
                    DataTable dt = new DataTable();
                    SqlCommand cmd_mail = new SqlCommand("SELECT * FROM tbl_gmail_settings", con);
                    SqlDataAdapter da = new SqlDataAdapter(cmd_mail);
                    da.Fill(dt);
                    SmtpClient smtp = new SmtpClient();
                    smtp.Host = dt.Rows[0]["Smtp_host"].ToString();
                    smtp.Port = Convert.ToInt32(dt.Rows[0]["Smtp_port"].ToString());
                    smtp.UseDefaultCredentials = false;
                    smtp.Credentials = new System.Net.NetworkCredential(dt.Rows[0]["Smtp_user"].ToString(), dt.Rows[0]["Smtp_pass"].ToString());
                    smtp.EnableSsl = true;
                    int i = 0;
                    //for (i = 0; i < MailToset.Length; i++)
                    //{
                    //    mail.To.Add(MailToset[i]);
                    //}
                    //i = 0;
                    //for (i = 0; i < CCset.Length; i++)
                    //{
                    //    mail.CC.Add(CCset[i]);
                    //}
                    //i = 0;
                    //for (i = 0; i < BCCset.Length; i++)
                    //{
                    //    mail.Bcc.Add(BCCset[i]);
                    //}
                    mail.To.Add("venkataprasad@titan.co.in");
                    mail.To.Add("annefebronia@titan.co.in");
                    mail.To.Add("pavithraashokan@titan.co.in");



                    //   mail.Bcc.Add("perennialdotnet@gmail.com");
                    mail.From = new MailAddress(dt.Rows[0]["Smtp_user"].ToString());
                    mail.Subject = "Daily Production Summary of " + strMachineName  + " Report on " + date.ToString("dd-MM-yyyy") + "";
                    mail.Body = htmlBody;
                    mail.IsBodyHtml = true;
                    if (attachmentFile != null)
                    {
                        Stream stream = new MemoryStream(attachmentFile);
                        mail.Attachments.Add(new Attachment(stream, fileName, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"));
                    }
                    smtp.Send(mail);
                    
                    strArray[0] = "1";
                    strArray[1] = "Emails Sent";
                    return  strArray;
                }
                catch (Exception ex)
                {
                    strArray[0] = "0";
                    strArray[1] = "Error while sending emails";
                    return strArray;
                    Console.WriteLine(ex.ToString());
                }
                finally
                {
                    con.Close();
                }
            }
        }
        //this method  is getting the linecode from selected dropdownlist
        //private string GetLineCode(string Id)
        //{
        //    return LineCodeList.Where(x => x.Value == Id).Select(x => x.Value).FirstOrDefault();
        //}
        //this method  is getting the plantcode from selected dropdownlist
        //private string GetPlantCode(string Id)
        //{
        //    return PlantCodeList.Where(x => x.Value == Id).Select(x => x.Value).FirstOrDefault();
        //}
        //this method  is getting the company code from selected dropdownlist
        //private string GetCompanyCode(string Id)
        //{
        //    return CompanyCodeList.Where(x => x.Value == Id).Select(x => x.Value).FirstOrDefault();
        //}
        //this method  is getting the station code from selected dropdownlist
        //private string GetStationCode(string Id)
        //{
        //    return StationCodeList.Where(x => x.Value == Id).Select(x => x.Value).FirstOrDefault();
        //}

        #region Return File Path For API

        //This method is used  to checking the Historical file based on  selected date,stationCode,stationCode is available or not if not available then call the  Generate Excel Path Report method

        public string DownloadExcelPath(ExcelReportViewModel model)
        {
            string MonthLabel = model.Date.ToString("MMMM_yyyy");
            string date = model.Date.ToString("dd");
            string day = model.Date.DayOfWeek.ToString();
            string lineCode = model.LineCode;
            string stationCode = model.StationCode;
            string fileName = MonthLabel + "_Day_" + date + "_" + lineCode + "_" + stationCode + "_Report.xlsx";
            var filePath = GetFilPathIfExist(fileName, lineCode, stationCode);
            string fileNameWithPath = "UploadFiles/Historical_ExcelFiles/" + lineCode + "/" + stationCode + "/" + fileName;
            var fileDownloadModel = new FileDownloadModel();
            if (filePath == true)
            {
                return fileNameWithPath;
            }
            else
            {
                return GenerateExcelPathReport(model);
            }
        }


        private bool GetFilPathIfExist(string name, string lineCode, string stationCode)
        {
            string path = Path.Combine(_hostingEnvironment.WebRootPath, "UploadFiles", "Historical_ExcelFiles", lineCode, stationCode, name);
            if (File.Exists(path))
            {
                return true;
            }
            return false;
        }

        private string GenerateExcelPathReport(ExcelReportViewModel model)
        {
            DateTime today = model.Date;
            var dat = today.ToString("yyyy-MM-dd");
            DateTime nxttoday = today.AddDays(1);
            var nxtdat = nxttoday.ToString("yyyy-MM-dd");
            string holidayname = "";
            string companyCode =model.CompanyCode;
            string companyName = model.CompanyName;
            string plantCode = model.PlantCode;
            string lineCode =model.LineCode;
            string stationCode = model.StationCode;
            
            using (SqlConnection connection = new SqlConnection(GetConnectionString()))
            {
                try
                {
                    SqlCommand holiDayCmd = new SqlCommand("SELECT [HolidayReason],[Date] FROM [dbo].[tbl_Holiday] where Date= @date and CompanyCode= @CompanyCode and PlantID= @PlantCode ", connection);
                    holiDayCmd.Parameters.AddWithValue("@date", dat);
                    holiDayCmd.Parameters.AddWithValue("@CompanyCode", companyCode);
                    holiDayCmd.Parameters.AddWithValue("@PlantCode", plantCode);
                    SqlDataAdapter holidayDa = new SqlDataAdapter(holiDayCmd);
                    DataTable holidayDt = new DataTable();
                    connection.Open();
                    holidayDa.Fill(holidayDt);
                    var first = new DateTime(today.Year, today.Month, 1);
                    var frstdaystr = first.ToString("yyyy-MM-dd");
                    if (dat == frstdaystr)
                    {
                        SqlCommand cmd555 = new SqlCommand("truncate table [dbo].[DailyReport_Time];", connection);
                        cmd555.ExecuteNonQuery();
                    }
                    else
                    {
                        SqlCommand cmd555 = new SqlCommand("delete from [dbo].[DailyReport_Time] where Date=@date;", connection);
                        cmd555.Parameters.AddWithValue("@date", dat);
                        cmd555.ExecuteNonQuery();
                    }

                    if (holidayDt.Rows.Count != 0)
                    {
                        holidayname = holidayDt.Rows[0][0].ToString();
                        return null;
                    }
                    else
                    {
                        SqlCommand assetCmd = new SqlCommand("select AssetID as Machine_Code, AssetName as MachineName,f.FunctionID as Line_code" +
                        //" from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID\", connection);
                        " from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID where f.FunctionID=@FunctionID AND AssetID=@AssetID", connection);
                        assetCmd.Parameters.AddWithValue("@FunctionID", lineCode);
                        assetCmd.Parameters.AddWithValue("@AssetID", stationCode);
                        SqlDataAdapter assetsDa = new SqlDataAdapter(assetCmd);
                        DataTable assetsDt = new DataTable();
                        assetsDa.Fill(assetsDt);
                        string[] name = new string[assetsDt.Rows.Count];
                        string[] line = new string[assetsDt.Rows.Count];
                        string[] machinecode = new string[assetsDt.Rows.Count];
                        for (int i = 0; i < assetsDt.Rows.Count; i++)
                        {
                            machinecode[i] = assetsDt.Rows[i][0].ToString();
                            name[i] = assetsDt.Rows[i][1].ToString();
                            line[i] = assetsDt.Rows[i][2].ToString();
                        }
                        var day = today.ToString("dddd");
                        string templatePath = Path.Combine(this._hostingEnvironment.WebRootPath, "UploadFiles", "Template", "Template_3.xlsx");
                        String MonthLabel = today.ToString("MMMM_yyyy");
                        String dayy = today.ToString("dd");
                        var excelFilePathWithLineCode = Path.Combine(ExcelFilePath, lineCode, stationCode);
                        if (!Directory.Exists(excelFilePathWithLineCode))
                        {
                            Directory.CreateDirectory(excelFilePathWithLineCode);
                        }
                        var datefile = model.Date.ToString("dd");
                        string fileName = MonthLabel + "_Day_" + datefile + "_" + lineCode + "_" + stationCode + "_Report.xlsx";

                        string nameFile = "UploadFiles/ExcelFiles" + "/" + lineCode + "/" + stationCode + "/" + fileName;

                        string filepath = Path.Combine(excelFilePathWithLineCode, fileName);
                        string wwwpath = "wwwroot/" + lineCode + stationCode + "/";
                        XLWorkbook workbook = new XLWorkbook(templatePath);

                        workbook.SaveAs(filepath);
                        for (int i = 0; i < assetsDt.Rows.Count; i++)
                        {
                            DataSet dataset = getDataSet(GetConnectionString(), line[i], today, companyCode, plantCode, machinecode[i]);
                            if (i == 0)
                            {
                                LoopAllMachines GF1 = new LoopAllMachines();
                                GF1.getData(GetConnectionString(), dataset, name[i], templatePath, filepath, line[i], dat);
                            }
                            GetExcelFile GF = new GetExcelFile();
                            GF.getData(GetConnectionString(), dataset, name[i], templatePath, filepath, line[i], companyCode, companyName, plantCode, machinecode[i], dat, i);
                        }
                        DirectoryInfo dirInfo = new DirectoryInfo(excelFilePathWithLineCode);
                        foreach (FileInfo excelFile in dirInfo.EnumerateFiles())
                        {
                            var historicalExcelFilePathWithLineCode = Path.Combine(Historical_ExcelFiles, lineCode, stationCode);
                            if (!Directory.Exists(historicalExcelFilePathWithLineCode))
                            {
                                Directory.CreateDirectory(historicalExcelFilePathWithLineCode);
                            }

                            if (!File.Exists(Path.Combine(historicalExcelFilePathWithLineCode, MonthLabel + "_Day_" + dayy + "_" + lineCode + "_" + stationCode + "_Report.xlsx")))
                            {
                                File.Copy(filepath, Path.Combine(historicalExcelFilePathWithLineCode, MonthLabel + "_Day_" + dayy + "_" + lineCode + "_" + stationCode + "_Report.xlsx"));
                                //excelFile.MoveTo(Historical_ExcelFiles);
                                //excelFile.Delete();
                            }
                        }
                        return nameFile;
                    }
                }
                catch (Exception ex)
                {
                    throw ex;
                }
                finally
                {
                    connection.Close();
                }
            }
        }
        #endregion

        #region Download and Send Mail

        //This method is used to gettting to,cc,bcc mail ids , download excel file and call the send mail methods
        public EmailResponse DownloadSendEmail(ExcelReportViewModel model, byte[] fileBytes, string fileName)
        {

            string[] strEmailRespose = new string[2];
 
            try
            {
                
                // byte[] fileBytes = DownloadExcel(model).FileBytes;
                string companyCode =model.CompanyCode;
                string plantCode =model.PlantCode;
                string lineCode =model.LineCode;
                string stationCode =model.StationCode;

                string[] toEmailAry = getemailDataSet(companyCode, plantCode, lineCode, stationCode);
                string[] ccEmailAry = getccDataSet(companyCode, plantCode, lineCode, stationCode);
                string[] bccEmailAry = getbccDataSet(companyCode, plantCode, lineCode, stationCode);
                if (toEmailAry.Length <= -1)
                {
                    return new EmailResponse() { IsSent = true, Message = "To email ids are not configured" };
                }
                if (ccEmailAry.Length <= -1)
                {
                    return new EmailResponse() { IsSent = true, Message = "cc email ids are not configured" };
                }
                string htmlBody = string.Empty;
                //if (fileBytes == null)
                //{
                //    htmlBody = "<p>Hello,</p>" +
                //                    "<p>Monthly Production Summary of VCTM Cycletime & Tool Life Report does not created.</p>" +
                //                    "<p>Possibly it can be holiday for date" + model.Date.ToString("dd-MM-yyyy") + "</p>";
                //}
                //else
                //{
                //    htmlBody = "<p>Hello,</p>" +
                //                    "<p>Monthly Production Summary of VCTM Cycletime & Tool Life Report</p>" +
                //                    "<p>Please find attached file</p>";
                //}

                //AddedBy/Vinayagamoorthi M/09/03/2023/Start

                string textBody = string.Empty;
                string strMachineName = "";

                var searchFile = Path.Combine(Historical_ExcelFiles, lineCode, stationCode);
                bool isFileExists = Directory.Exists(searchFile);

                if (isFileExists == false)
                {
                    Directory.CreateDirectory(searchFile);
                }

                string[] files = Directory.GetFiles(searchFile, fileName);


                if (fileBytes == null)
                {
                    htmlBody = "<p>Production count of <b>" + plantCode + " - " + lineCode + ":" + model.Date.ToString("yyyy-MM-dd") + "</b></p>" +
                                    "<p><b>Day Start Date:" + model.Date.ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>" +
                                    "<p><b>Day End Date:" + model.Date.AddDays(1).ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>";

                }
                else
                {

                    htmlBody = "<p>Production count of <b>" + plantCode + " - " + lineCode + ":" + model.Date.ToString("yyyy-MM-dd") + "</b></p>" +
                                    "<p><b>Day Start Date:" + model.Date.ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>" +
                                    "<p><b>Day End Date:" + model.Date.AddDays(1).ToString("yyyy-MM-dd") + " 08:15:00 AM" + "</b></p>";

                    textBody = " <table bordercolor='#00A8FC' border=" + 1 + " cellpadding=" + 0 + " cellspacing=" + 0 + " width = " + 900 + "><tr  bordercolor='#00A8FC' bgcolor='#00A8FC'>" +
                        "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#FFFFFF'><b>Machine Name</b></td> " +
                        "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#FFFFFF'><b>Variant Name</b></td>" +
                        "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#FFFFFF'><b>Actual Ok Parts</b></td>" +
                        "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#FFFFFF'><b>Actual NOk Parts</b></td>" +
                        "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#FFFFFF'><b>Rejection %</b></td>" +
                        "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#FFFFFF'><b>UpTime (In Mins/in %)</b></td>" +
                        "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#FFFFFF'><b>Downtime (In Mins/in %)</b></td>" +
                        "</tr>";

                    string[] name;
                    string[] line;
                    string[] machinecode;
                    using (SqlConnection connection = new SqlConnection(GetConnectionString()))
                    {
                        try
                        {
                            SqlCommand assetCmd = new SqlCommand("select AssetID as Machine_Code, AssetName as MachineName,f.FunctionID as Line_code" +
                            " from tbl_Assets a inner join tbl_function f on a.FunctionName=f.FunctionID where f.FunctionID=@FunctionID AND AssetID=@AssetID", connection);
                            assetCmd.Parameters.AddWithValue("@FunctionID", lineCode);
                            assetCmd.Parameters.AddWithValue("@AssetID", stationCode);
                            SqlDataAdapter assetsDa = new SqlDataAdapter(assetCmd);
                            DataTable assetsDt = new DataTable();
                            connection.Open();
                            assetsDa.Fill(assetsDt);
                            name = new string[assetsDt.Rows.Count];
                            line = new string[assetsDt.Rows.Count];
                            machinecode = new string[assetsDt.Rows.Count];
                            for (int i = 0; i < assetsDt.Rows.Count; i++)
                            {
                                machinecode[i] = assetsDt.Rows[i][0].ToString();
                                name[i] = assetsDt.Rows[i][1].ToString();
                                line[i] = assetsDt.Rows[i][2].ToString();
                            }

                        }
                        catch (Exception ex)
                        {
                            throw ex;
                        }
                        finally

                        {
                            connection.Close();
                        }
                    }


                    DataSet dataset = getDataSet(GetConnectionString(), lineCode, model.Date, companyCode, plantCode, stationCode);

                    if (dataset.Tables[1].Rows.Count > 0)
                    {

                        float strRejection = 0;
                        for (int i = 0; i < name.Length; i++)
                        {
                            if (stationCode == machinecode[i] && lineCode == line[i])
                            {
                                strMachineName = name[i].ToString();
                            }
                        }

                        foreach (DataRow Row in dataset.Tables[1].Rows)
                        {

                            if ((Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"])) != 0)
                            {
                                strRejection = ((float)Math.Round((((float)(Convert.ToInt32(Row["actual_nok"])) / (float)(Convert.ToInt32(Row["ActualProduction"]) + Convert.ToInt32(Row["actual_nok"]))) * 100) * 100f) / 100f);
                            }
                            else
                            {
                                strRejection = 0;
                            }



                            textBody += "<tr>" +
                                "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'>" + strMachineName + "</td>" +
                                "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'> " + Row["VariantCode"].ToString() + "</td> " +
                                "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> " + Row["ActualProduction"].ToString() + "</td> " +
                                "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> " + Row["actual_nok"].ToString() + "</td> " +
                                "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#1A0303'> " + strRejection.ToString() + "</td> " +
                                "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> " + Row["UpTime(min)"].ToString() + " Mins / " + Row["UPTime%"].ToString() + " %" + " </td> " +
                                "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> " + Row["DownTime(Min)"].ToString() + " Mins / " + Row["DownTime%"].ToString() + " %" + " </td> " +
                                "</tr>";
                        }
                        textBody += "</table>";
                    }
                    else
                    {
                        for (int i = 0; i < name.Length; i++)
                        {

                            string dayname = "";
                            dayname = model.Date.ToString("dddd").ToString();
                            if (dayname == "Sunday")
                            {
                                textBody += "<tr>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'>" + name[i].ToString() + "</td>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'> --Sunday--Holiday-- </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "</tr>";
                            }
                            else
                            {
                                textBody += "<tr>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'>" + name[i].ToString() + "</td>" +
                                            "<td style='padding:10px; height:24px; width=120px;text-align:center;color:#1A0303'> --Data Not Logged-- </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=125px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=110px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "<td style='padding:10px; height:24px; width=130px;text-align:center;color:#1A0303'> - </td> " +
                                            "</tr>";
                            }


                        }
                        textBody += "</table>";

                    }
                

                    textBody += "<p>For more dteails refer the attachment!</p>" +
                                    "<p>Refer the portal for more info <u style='color:#00A8FC'>" + "click  to login" + "</u></p>" +
                                    "<p>***Mail generated from TEAL IIOT Portal Email App services***</p>";

                    htmlBody += textBody;
                }

                strEmailRespose =  SendEmail(htmlBody, toEmailAry, ccEmailAry, bccEmailAry, model.Date, fileBytes, fileName, strMachineName);
                return new EmailResponse() { IsSent = Convert.ToBoolean(Convert.ToInt32 (strEmailRespose[0])), Message = strEmailRespose[1], ToSent = toEmailAry, CCSent = ccEmailAry,BCCSent =bccEmailAry };
            }
            catch (Exception ex)
            {
                if (strEmailRespose[0] == null)
                {
                    return new EmailResponse() { IsSent = false, Message = ex.ToString() };
                }
                else
                {
                    return new EmailResponse() { IsSent = Convert.ToBoolean(Convert.ToInt32(strEmailRespose[0])), Message = strEmailRespose[1].ToString() };
                }

            }
            finally
            {

            }
        }
        #endregion

        #region Return File From File Path
        //this method used to Get the File name From File Path
        public FileDownloadModel GetFileFromFilePath(string filePathName)
        {
            FileDownloadModel model = new FileDownloadModel();
            string path = Path.Combine(_hostingEnvironment.WebRootPath, filePathName);
            if (File.Exists(path))
            {
                model.FileBytes = File.ReadAllBytes(path);
                model.FileName = Path.GetFileName(path);

            }
            return model;
        }
        //this method used to Get the download excel from File Path
        public FileDownloadModel DownloadExcelFileFromPath(string filePath)
        {
            var fileDownloadModel = GetFileFromFilePath(filePath);
            if (fileDownloadModel.FileBytes != null)
            {
                fileDownloadModel.FileBytes = fileDownloadModel.FileBytes;
                fileDownloadModel.FileName = fileDownloadModel.FileName;
            }
            return fileDownloadModel;
        }
        #endregion

    }
}