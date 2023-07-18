using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data.SqlClient;
using System.Data;
using System.Linq;
using System.Web;

namespace LessonLearntPortalWeb.Repository
{
    public class LoopAllMachines
    {


        public void getData(string connStr, DataSet ds, string machinecode, string path, string filepath, string linecode, string date)
        {
            DataSet ds1 = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();

                    Console.WriteLine("Data required for excel has been collected ");

                    ///variant list of production qty variant-wise and day-wise - TABLE 7
                    ds1.Tables.Add(ds.Tables[7].Copy());

                    UploadExcelProduction(ds1, machinecode, connStr, path, filepath, date);

                    //ExportDataSetToExcel(ds);
                    Console.WriteLine("Excel Chart has been generated");


                }
                catch (SqlException ex)
                {

                }
                catch (Exception e)
                {

                }
            }

        }


        public static void UploadExcelProduction(DataSet ds, string machinecode, String connStr, String path, String filepath, string date)
        {

            try
            {

            
                date = Convert.ToDateTime(date).AddDays(-1).ToString("yyyy-MM-dd");
                string datetoaddinTime = Convert.ToDateTime(date).AddDays(-1).ToString("yyyy-MM-dd");

                DataTable dt10 = new DataTable();

                dt10 = ds.Tables[0];

                //Started reading the Excel file.  
                using (XLWorkbook workbook = new XLWorkbook(filepath))
                {


                    ////variant entering in cummulative production qty sheet
                    IXLWorksheet ws8 = workbook.Worksheet(2);

                    if (dt10.Rows.Count > 0)
                    {
                        int aa2 = 6;
                        // Adding DataRows.
                        for (int i = 0; i < dt10.Rows.Count; i++)
                        {

                            ws8.Cell("A" + (aa2)).Value = dt10.Rows[i][0].ToString();

                            aa2++;
                        }
                    }

                    DateTime dat = Convert.ToDateTime(date).AddDays(-1);

                    var dates = new List<DateTime>();

                    var firstDayOfMonth = new DateTime(dat.Year, dat.Month, 1);
                    var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

                    var NoOfMachine = 5;

                    for (var dt33 = dat; dt33 >= firstDayOfMonth; dt33 = dt33.AddDays(-1))
                    {
                        dates.Add(dt33);
                    }

                    //////DAY-WISE production qty
                    //IXLWorksheet ws11 = workbook.Worksheet(3);

                    //var datecolumnName = "H";

                    //for (int i12 = 0; i12 < dates.Count; i12++)
                    //{
                    //    ws11.Cell(datecolumnName + "4").Value = dates[i12].ToString();
                    //    ws11.Cell(datecolumnName + "4").Style.Font.Bold = true;
                    //    ws11.Cell(datecolumnName + "4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                    //    int sum111 = sum(datecolumnName);

                    //    datecolumnName = calculation(sum111 + 5);


                    //}

                    workbook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Style.Font.Bold = true;
                    workbook.Save();
                    //workbook.SaveAs(filepath);
                }

            }
            catch (Exception ex)
            {
                throw ex;
            }
            finally
            {

            }
        }
    }
}