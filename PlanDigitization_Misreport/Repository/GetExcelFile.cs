using ClosedXML.Excel;
using DocumentFormat.OpenXml.Vml;
using DocumentFormat.OpenXml.Wordprocessing;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;

namespace LessonLearntPortalWeb.Repository
{
    public class GetExcelFile
    {
        public void getData(string connStr, DataSet ds, string machinename, string path, string filepath, string linecode, string CompanyCode,string companyName, string PlantCode, string machinecode, string date, int iteration)
        {
            DataSet ds1 = new DataSet();
            using (SqlConnection con = new SqlConnection(connStr))
            {
                try
                {
                    con.Open();
                    Console.WriteLine("Data required for excel has been collected ");
                    ///Hourly Tracker - Table 0
                    ds1.Tables.Add(ds.Tables[0].Copy());

                    ///Mailbody - TABLE 1
                    ds1.Tables.Add(ds.Tables[1].Copy());

                    ///variant-wise Utilisation - TABLE 2
                    ds1.Tables.Add(ds.Tables[2].Copy());

                    ///diagnostic details - TABLE 3
                    ds1.Tables.Add(ds.Tables[3].Copy());

                    ///Top 10 Error data - TABLE 4
                    ds1.Tables.Add(ds.Tables[4].Copy());

                    ///Cummulative production qty variant-wise - TABLE 5
                    ds1.Tables.Add(ds.Tables[5].Copy());

                    ///variant wise production qty for month - day-wise - TABLE 6
                    ds1.Tables.Add(ds.Tables[6].Copy());

                    ///variant list of production qty - TABLE 7
                    ds1.Tables.Add(ds.Tables[7].Copy());

                    ///variant list rejection - TABLE 8
                    ds1.Tables.Add(ds.Tables[8].Copy());

                    ///Tool life - TABLE 9
                    ds1.Tables.Add(ds.Tables[9].Copy());

                    //Cycletime
                    ds1.Tables.Add(ds.Tables[10].Copy());

                    //Batchwise Hourly Tracker
                    ds1.Tables.Add(ds.Tables[11].Copy());

                    UploadExcelProduction(ds1, machinename, connStr, path, filepath, linecode, CompanyCode, companyName, PlantCode, machinecode, date, iteration);
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
        public static void UploadExcelProduction(DataSet ds, string machinename, String connStr, String path, String filepath, string linecode, string CompanyCode,string companyName, string PlantCode, string machinecode, string date, int iteration)
        {
            try
            {
                DataTable dt1 = new DataTable();
                DataTable dt2 = new DataTable();
                DataTable dt3 = new DataTable();
                DataTable dt4 = new DataTable();
                DataTable dt5 = new DataTable();
                DataTable dt6 = new DataTable();
                DataTable dt7 = new DataTable();
                DataTable dt8 = new DataTable();
                DataTable dt9 = new DataTable();
                DataTable dt10 = new DataTable();
                DataTable dt11 = new DataTable();
                DataTable dt12 = new DataTable();
                DataTable dt13 = new DataTable();


                dt1 = ds.Tables[0];     //hourly
                dt2 = ds.Tables[1];     //mail
                dt3 = ds.Tables[2];     //var util
                dt4 = ds.Tables[3];     //diag
                dt5 = ds.Tables[4];     //alarm
                dt6 = ds.Tables[5];     //cumm
                dt7 = ds.Tables[6];     //month
                dt8 = ds.Tables[7];     //prod
                dt9 = ds.Tables[8];     //variant rej
                dt10 = ds.Tables[9];    //tool 
                dt11 = ds.Tables[10];  //Cycletime
                dt12 = ds.Tables[11]; //Batchwise Hourly Tracker




                //Started reading the Excel file.  
                using (XLWorkbook workbook = new XLWorkbook(filepath))
                {
                    int sheetno = 0;
                    var column1 = "";
                    var column2 = "";
                    var column3 = "";
                    var columnNam = "";
                    int aa1 = 0;

                    //if (machinecode == "M1")
                    //{
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    sheetno = 6;
                    //sheetno = 5;
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End


                    column1 = "B";
                        column2 = "C";
                        column3 = "D";
                        columnNam = "B";
                        aa1 = 3;
                    //}


                    ////production summary
                  //  IXLWorksheet ws = workbook.Worksheet(sheetno);
                    //Console.WriteLine(ws + " : " + ws.Cell("F" + 11).Value);
                    //ws.Cell("F" + 11).Value = date;
                    //if (dt2.Rows.Count > 0)
                    //{
                        // Adding HeaderRow.
                        //Console.WriteLine(ws.Cell("A" + 2).Value);
                        //Console.WriteLine(ws.Cell("B" + 2).Value);
                        //Console.WriteLine(ws.Cell("C" + 2).Value);
                        //ws.Cell("A" + 2).Value = dt2.Columns[0].ColumnName;
                        //ws.Cell("B" + 2).Value = dt2.Columns[1].ColumnName;
                        //ws.Cell("C" + 2).Value = dt2.Columns[2].ColumnName;

                        // Adding DataRows.
                        //for (int i = 0; i < dt2.Rows.Count; i++)
                        //{
                        //    ws.Cell("A" + (i + 3)).Value = dt2.Rows[i][0].ToString();
                        //    ws.Cell("B" + (i + 3)).Value = Convert.ToDecimal(dt2.Rows[i][1]);
                        //    ws.Cell("C" + (i + 3)).Value = Convert.ToDecimal(dt2.Rows[i][2]);

                        //}
                   // }
                    //(XLCellValue)
                    //(XLCellValue)
                    //(XLCellValue)
                 //   workbook.Save();


                    ////Top 10 rejections
                    IXLWorksheet ws13 = workbook.Worksheet(sheetno + 1);
                    if (dt9.Rows.Count > 0)
                    {
                        int k = 4;

                        // Adding DataRows.
                        for (int i = 0; i < dt9.Rows.Count; i++)
                        {

                            DataTable data = new DataTable();

                            SqlConnection con = new SqlConnection(connStr);
                            SqlCommand cmd = new SqlCommand(@"select top 10 b.variant_code,
								case when b.Reject_reason=''  then '0' else b.Reject_reason end,count(b.Reject_reason) as total 
								,a.RejectionDescription
								from [dbo].[tbl_Product_Reject_reason] b 
								join tbl_Rejection a on b.Reject_Reason=a.RejectionCode
								where b.Date =@date and b.line_code =@line and b.CompanyCode=@company 
								and b.PlantCode=@plant and b.[Machine_code]=@machine and b.Variant_code=@variant
								group by b.variant_code, b.Reject_reason ,a.RejectionDescription
								order by total desc", con);

                            cmd.Parameters.AddWithValue("@date", date);
                            cmd.Parameters.AddWithValue("@line", linecode);
                            cmd.Parameters.AddWithValue("@company", CompanyCode);
                            cmd.Parameters.AddWithValue("@plant", PlantCode);
                            cmd.Parameters.AddWithValue("@machine", machinecode);
                            cmd.Parameters.AddWithValue("@variant", dt9.Rows[i][0]);
                            SqlDataAdapter da = new SqlDataAdapter(cmd);
                            da.Fill(data);

                            if (data.Rows.Count > 0)
                            {
                                for (int j = 0; j < data.Rows.Count; j++)
                                {

                                    ws13.Cell("A" + (j + k)).Value = j + 1;
                                    ws13.Cell("B" + (j + k)).Value = data.Rows[j][0].ToString();
                                    ws13.Cell("C" + (j + k)).Value = data.Rows[j][1].ToString();
                                    ws13.Cell("D" + (j + k)).Value = Convert.ToDecimal(data.Rows[j][2]);
                                    ws13.Cell("E" + (j + k)).Value = data.Rows[j][3].ToString();
                                }
                                k += 12;
                            }
                        }
                    }
                    workbook.Save();

                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start

                    //////Variant-wise Utilisation summary
                    //IXLWorksheet ws4 = workbook.Worksheet(3);

                    //if (dt3.Rows.Count > 0)
                    //{
                    //    ws4.Cell("B" + 2).Value = "Production On";
                    //    var Varaint_range = ws4.Range("A4", "I13");
                    //    int aa = Varaint_range.FirstRow().RowNumber();
                    //    // Adding DataRows
                    //    for (int i = 0; i < dt3.Rows.Count; i++)
                    //    {


                    //        ws4.Cell("A" + (aa)).Value = dt3.Rows[i][0].ToString();
                    //        ws4.Cell("B" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][1]);
                    //        ws4.Cell("C" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][2]);
                    //        ws4.Cell("D" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][3]);
                    //        ws4.Cell("E" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][4]);
                    //        ws4.Cell("F" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][5]);
                    //        ws4.Cell("G" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][6]);
                    //        ws4.Cell("H" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][7]);
                    //        ws4.Cell("I" + (aa)).Value = Convert.ToDecimal(dt3.Rows[i][8]);

                    //        aa += 1;

                    //    }
                    //}
                    //else if (dt3.Rows.Count == 0)
                    //{
                    //    ws4.Cell("B" + 2).Value = "No Production";
                    //}

                    //workbook.Save();

                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End

                    if (iteration == 0)
                    {
                        ////Diagnostic details

                        //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                        IXLWorksheet ws5 = workbook.Worksheet(4);
                        //IXLWorksheet ws5 = workbook.Worksheet(3);
                        //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End

                        if (dt4.Rows.Count > 0)
                        {

                            // Adding DataRows.
                            for (int i = 0; i < dt4.Rows.Count; i++)
                            {
                                int aa = ws5.LastRowUsed().RowNumber() + 1;

                                //int row = ws5.LastRowUsed().RowNumber();

                                ws5.Cell("A" + (aa)).Value = dt4.Rows[i][0].ToString();
                                ws5.Cell("B" + (aa)).Value = dt4.Rows[i][1].ToString();
                                ws5.Cell("C" + (aa)).Value = dt4.Rows[i][2].ToString();
                                ws5.Cell("D" + (aa)).Value = dt4.Rows[i][3].ToString();
                                ws5.Cell("E" + (aa)).Value = dt4.Rows[i][4].ToString();
                                ws5.Cell("F" + (aa)).Value = dt4.Rows[i][4].ToString();

                            }
                        }
                        IXLCell firstcelldiag = ws5.Cell("A3");

                        IXLCell lastcelldiag = ws5.LastCellUsed();

                        // the range for which you want to add a table style
                        var rngTable1diag = ws5.Range(firstcelldiag, lastcelldiag);

                        rngTable1diag.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                        rngTable1diag.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                        rngTable1diag.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        rngTable1diag.Style.Border.RightBorder = XLBorderStyleValues.Thin;



                    }

                    workbook.Save();

                    ////Stationwise Top 10 Errors

                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    IXLWorksheet ws7 = workbook.Worksheet(5);
                    //IXLWorksheet ws7 = workbook.Worksheet(4);
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End


                    if (dt5.Rows.Count > 0)
                    {

                        // Adding DataRows.
                        for (int i = 0; i < dt5.Rows.Count; i++)
                        {
                            ws7.Cell("A" + (i + 5)).Value = i + 1;
                            ws7.Cell(column1 + (i + 5)).Value = dt5.Rows[i][0].ToString();
                            ws7.Cell(column2 + (i + 5)).Value = Convert.ToDecimal(dt5.Rows[i][1]);
                            ws7.Cell(column3 + (i + 5)).Value = Convert.ToDecimal(dt5.Rows[i][2]);
                        }
                    }

                    workbook.Save();



                    //////variant entering in cummulative production qty sheet
                    //IXLWorksheet ws8 = workbook.Worksheet(3);

                    //if (dt10.Rows.Count > 0)
                    //{
                    //    int aa2 = 6;
                    //    // Adding DataRows.
                    //    for (int i = 0; i < dt10.Rows.Count; i++)
                    //    {

                    //        ws8.Cell("A" + (aa2)).Value = dt10.Rows[i][0];

                    //        aa2++;
                    //    }
                    //}


                    ////cummulative production qty
                    IXLWorksheet ws9 = workbook.Worksheet(2);

                    if (dt6.Rows.Count > 0)
                    {

                        ////--------for cummulative data entering---------////
                        for (int j = 0; j < dt6.Rows.Count; j++)
                        {
                            var aa = dt6.Rows[j][0].ToString();


                            var dd = ws9.CellsUsed(x => x.GetString() == aa);


                            ////to find the respective cell in a n excel
                            //var dd = ws4.Search(aa, CompareOptions.OrdinalIgnoreCase);

                            search(dd, ws9, dt6.Rows[j][1].ToString(), columnNam);

                        }

                    }

                    workbook.Save();

                    ////Hourly Tracker
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    IXLWorksheet ws32 = workbook.Worksheet(9);
                    //IXLWorksheet ws32 = workbook.Worksheet(8);
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End


                    if (dt1.Rows.Count > 0)
                    {
                        int aa12 = 3;
                        string columnNamadd4 = "A" + aa12;


                        // Adding DataRows.
                        for (int i = 0; i < dt1.Rows.Count; i++)
                        {
                            ws32.Cell("A" + (aa12)).Value = dt1.Rows[i][0].ToString();
                            ws32.Cell("B" + (aa12)).Value = dt1.Rows[i][1].ToString();
                            ws32.Cell("C" + (aa12)).Value = dt1.Rows[i][2].ToString();
                            ws32.Cell("D" + (aa12)).Value = dt1.Rows[i][3].ToString();
                            ws32.Cell("E" + (aa12)).Value = dt1.Rows[i][4].ToString();

                            
                            if (dt1.Rows[i][5].ToString() != "")
                            {
                                ws32.Cell("F" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][5]);
                            }
                            if (dt1.Rows[i][6].ToString() != "")
                            {
                                ws32.Cell("G" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][6]);
                            }
                            if (dt1.Rows[i][7].ToString() != "")
                            {
                                ws32.Cell("H" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][7]);
                            }
                            if (dt1.Rows[i][8].ToString() != "")
                            {
                                ws32.Cell("I" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][8]);
                            }
                            if (dt1.Rows[i][9].ToString() != "")
                            {
                                ws32.Cell("J" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][9]);
                            }

                            ws32.Cell("K" + (aa12)).Value = dt1.Rows[i][10].ToString();

                            if (dt1.Rows[i][11].ToString() != "")
                            {
                                ws32.Cell("L" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][11]);
                            }
                            if (dt1.Rows[i][12].ToString() != "")
                            {
                                ws32.Cell("M" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][12]);
                            }
                            if (dt1.Rows[i][13].ToString() != "")
                            {
                                ws32.Cell("N" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][13]);
                            }
                            if (dt1.Rows[i][14].ToString() != "")
                            {
                                ws32.Cell("O" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][14]);
                            }
                            if (dt1.Rows[i][15].ToString() != "")
                            {
                                ws32.Cell("P" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][15]);
                            }
                            if (dt1.Rows[i][16].ToString() != "")
                            {
                                ws32.Cell("Q" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][16]);
                            }
                            if (dt1.Rows[i][17].ToString() != "")
                            {
                                ws32.Cell("R" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][17]);
                            }
                            if (dt1.Rows[i][18].ToString() != "")
                            {
                                ws32.Cell("S" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][18]);
                            }
                            if (dt1.Rows[i][19].ToString() != "")
                            {
                                ws32.Cell("T" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][19]);
                            }
                            if (dt1.Rows[i][20].ToString() != "")
                            {
                                ws32.Cell("U" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][20]);
                            }
                            if (dt1.Rows[i][21].ToString() != "")
                            {
                                ws32.Cell("V" + (aa12)).Value = Convert.ToDecimal(dt1.Rows[i][21]);
                            }
                            //ws32.Cell("W" + (aa12)).Value = dt1.Rows[i][22]; // add new
                            //ws32.Cell("X" + (aa12)).Value = dt1.Rows[i][23]; /// add new 

                            aa12++;
                        }

                        aa12--;

                        IXLCell firstcell4 = ws32.Cell(columnNamadd4);

                        IXLCell lastcell4 = ws32.LastCellUsed();

                        // the range for which you want to add a table style
                        var rngTable4 = ws32.Range(firstcell4, lastcell4);

                        rngTable4.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                        rngTable4.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                        rngTable4.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        rngTable4.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    }

                    workbook.Save();

                    //-----Tool Life sheet -------
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    IXLWorksheet ws36 = workbook.Worksheet(10);
                    //IXLWorksheet ws36 = workbook.Worksheet(9);
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End

                    if (dt10.Rows.Count > 0)
                    {

                        int aa2 = 6;

                        for (int i = 0; i < dt10.Rows.Count; i++)
                        {

                            ws36.Cell("A" + (aa2)).Value = i + 1;
                            ws36.Cell("B" + (aa2)).Value = dt10.Rows[i][0].ToString();
                            ws36.Cell("C" + (aa2)).Value = dt10.Rows[i][1].ToString();
                            ws36.Cell("D" + (aa2)).Value = dt10.Rows[i][2].ToString();
                            ws36.Cell("E" + (aa2)).Value = dt10.Rows[i][3].ToString();
                            ws36.Cell("F" + (aa2)).Value = dt10.Rows[i][4].ToString();
                            ws36.Cell("G" + (aa2)).Value = dt10.Rows[i][5].ToString();
                            ws36.Cell("H" + (aa2)).Value = dt10.Rows[i][6].ToString();
                            ws36.Cell("I" + (aa2)).Value = dt10.Rows[i][7].ToString();
                            ws36.Cell("J" + (aa2)).Value = dt10.Rows[i][8].ToString();
                            ws36.Cell("K" + (aa2)).Value = dt10.Rows[i][9].ToString();
                            ws36.Cell("L" + (aa2)).Value = dt10.Rows[i][13].ToString();

                            aa2++;
                        }

                        IXLCell firstcelldiag = ws36.Cell("A5");

                        IXLCell lastcelldiag = ws36.LastCellUsed();

                        // the range for which you want to add a table style
                        var rngTable1diag = ws36.Range(firstcelldiag, lastcelldiag);

                        rngTable1diag.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                        rngTable1diag.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                        rngTable1diag.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        rngTable1diag.Style.Border.RightBorder = XLBorderStyleValues.Thin;

                    }
                    else
                    {

                        ws36.Cell("F" + 6).Value = "Data Not Available";
                        ws36.Cell("F" + 6).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5a742");
                    }

                    workbook.Save();

                    //-----Cycletime sheet -------

                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    IXLWorksheet wscy = workbook.Worksheet(11);
                    //IXLWorksheet wscy = workbook.Worksheet(10);
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End



                    if (dt11.Rows.Count > 0)
                    {
                        int aa2 = 3;

                        for (int j = 0; j < dt11.Rows.Count; j++)
                        {

                            wscy.Cell("A" + (aa2)).Value = "'" + dt11.Rows[j][0].ToString();


                            for (int i = 1; i < dt11.Columns.Count; i++)
                            {
                                if (dt11.Columns[i].ColumnName == "V1")
                                {
                                    if (dt11.Rows[j][i].ToString() != "")
                                    {
                                        wscy.Cell("B" + (aa2)).Value = Convert.ToDecimal(dt11.Rows[j][i]);
                                    }
                                }
                                if (dt11.Columns[i].ColumnName == "V2")
                                {
                                    if (dt11.Rows[j][i].ToString() != "")
                                    {
                                        wscy.Cell("C" + (aa2)).Value = Convert.ToDecimal(dt11.Rows[j][i]);
                                    }
                                }
                                if (dt11.Columns[i].ColumnName == "V3")
                                {
                                    if (dt11.Rows[j][i].ToString() != "")
                                    {
                                        wscy.Cell("D" + (aa2)).Value = Convert.ToDecimal(dt11.Rows[j][i]);
                                    }
                                }
                                if (dt11.Columns[i].ColumnName == "V4")
                                {
                                    if (dt11.Rows[j][i].ToString() != "")
                                    {
                                        wscy.Cell("E" + (aa2)).Value = Convert.ToDecimal(dt11.Rows[j][i]);
                                    }
                                }
                                if (dt11.Columns[i].ColumnName == "V5")
                                {
                                    if (dt11.Rows[j][i].ToString() != "")
                                    {
                                        wscy.Cell("F" + (aa2)).Value = Convert.ToDecimal(dt11.Rows[j][i]);
                                    }
                                    
                                }
                                if (dt11.Columns[i].ColumnName == "V6")
                                {
                                    if (dt11.Rows[j][i].ToString() != "")
                                    {
                                        wscy.Cell("G" + (aa2)).Value = Convert.ToDecimal(dt11.Rows[j][i]);
                                    }
                                }
                                
                            }
                            aa2++;
                        }

                    }
                    else
                    {
                        wscy.Cell("A" + 3).Value = "Data Not Available";
                        wscy.Cell("A" + 3).Style.Fill.BackgroundColor = XLColor.FromHtml("#f5a742");
                    }

                    workbook.Save();


                    ////Batchwise Hourly Tracker

                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    IXLWorksheet ws_hourly = workbook.Worksheet(12);
                    //IXLWorksheet ws_hourly = workbook.Worksheet(11);
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End

                    if (dt12.Rows.Count > 0)
                    {
                        int aa12 = 3;
                        string columnNamadd4 = "A" + aa12;


                        // Adding DataRows.
                        for (int i = 0; i < dt12.Rows.Count; i++)
                        {
                            ws_hourly.Cell("A" + (aa12)).Value = dt12.Rows[i][0].ToString();
                            ws_hourly.Cell("B" + (aa12)).Value = dt12.Rows[i][1].ToString();
                            ws_hourly.Cell("C" + (aa12)).Value = dt12.Rows[i][2].ToString();
                            ws_hourly.Cell("D" + (aa12)).Value = dt12.Rows[i][3].ToString();
                            ws_hourly.Cell("E" + (aa12)).Value = dt12.Rows[i][4].ToString();

                            if (dt12.Rows[i][5].ToString() != "")
                            {
                                ws_hourly.Cell("F" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][5]);
                            }
                            if (dt12.Rows[i][6].ToString() != "")
                            {
                                ws_hourly.Cell("G" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][6]);
                            }
                            if (dt12.Rows[i][7].ToString() != "")
                            {
                                ws_hourly.Cell("H" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][7]);
                            }
                            if (dt12.Rows[i][8].ToString() != "")
                            {
                                ws_hourly.Cell("I" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][8]);
                            }
                            if (dt12.Rows[i][9].ToString() != "")
                            {
                                ws_hourly.Cell("J" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][9]);
                            }
                            
                            ws_hourly.Cell("K" + (aa12)).Value = dt12.Rows[i][10].ToString();
                            
                            if (dt12.Rows[i][11].ToString() != "")
                            {
                                ws_hourly.Cell("L" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][11]);
                            }
                            if (dt12.Rows[i][12].ToString() != "")
                            {
                                ws_hourly.Cell("M" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][12]);
                            }
                            if (dt12.Rows[i][13].ToString() != "")
                            {
                                ws_hourly.Cell("N" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][13]);
                            }
                            if (dt12.Rows[i][14].ToString() != "")
                            {
                                ws_hourly.Cell("O" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][14]);
                            }
                            if (dt12.Rows[i][15].ToString() != "")
                            {
                                ws_hourly.Cell("P" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][15]);
                            }
                            if (dt12.Rows[i][16].ToString() != "")
                            {
                                ws_hourly.Cell("Q" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][16]);
                            }
                            if (dt12.Rows[i][17].ToString() != "")
                            {
                                ws_hourly.Cell("R" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][17]);
                            }
                            if (dt12.Rows[i][18].ToString() != "")
                            {
                                ws_hourly.Cell("S" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][18]);
                            }
                            if (dt12.Rows[i][19].ToString() != "")
                            {
                                ws_hourly.Cell("T" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][19]);
                            }
                            if (dt12.Rows[i][20].ToString() != "")
                            {
                                ws_hourly.Cell("U" + (aa12)).Value = Convert.ToDecimal(dt12.Rows[i][20]);
                            }
                            aa12++;
                        }

                        aa12--;

                        IXLCell firstcell4 = ws32.Cell(columnNamadd4);

                        IXLCell lastcell4 = ws32.LastCellUsed();

                        // the range for which you want to add a table style
                        var rngTable4 = ws32.Range(firstcell4, lastcell4);

                        rngTable4.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                        rngTable4.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                        rngTable4.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                        rngTable4.Style.Border.RightBorder = XLBorderStyleValues.Thin;
                    }

                    workbook.Save();



                    //--- fetching cell details from sheet 1 for Day-wise production ---//


                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/Start
                    IXLWorksheet wss = workbook.Worksheet(8);
                    //IXLWorksheet wss = workbook.Worksheet(7);
                    //AddedBy/Vinayagamoorthi M/Perennial Innovative Solutions/19/05/2023/End

                    var row = wss.Row(1);

                    var cell = row.Cell(2);

                    string value = cell.GetValue<string>();

                    var columnName = value.ToUpperInvariant();

                    int sum1 = sum(columnName);

                  //  DateTime dat =  Convert.ToDateTime(date).AddDays(-1);
                    DateTime dat = Convert.ToDateTime(date);
                    
                    
                    var dates = new List<DateTime>();

                    var firstDayOfMonth = new DateTime(dat.Year, dat.Month, 1);
                    var lastDayOfMonth = firstDayOfMonth.AddMonths(1).AddDays(-1);

                    var NoOfMachine = 1;

                    for (var dt33 = dat; dt33 >= firstDayOfMonth; dt33 = dt33.AddDays(-1))
                    {
                        dates.Add(dt33);
                    }

                    workbook.Save();


                    //---------------------------------------------------
                    //------------------------------------------------------

                    columnName = "G";
                    ////DAY-WISE production qty
                    IXLWorksheet ws11 = workbook.Worksheet(2);

                    int sum11 = sum(columnName);

                    var columnNamadd = calculation(sum11);
                    var columnNamadd1 = calculation(sum11);
                    //dates loop

                    for (int i11 = 0; i11 < dates.Count; i11++)
                    {

                        Console.WriteLine("------" + dates[i11] + "------");

                        Console.WriteLine("Started");

                        string holidayname = "";


                        SqlConnection con = new SqlConnection(connStr);
                        SqlCommand cmd = new SqlCommand("SELECT [HolidayReason],[Date] FROM [dbo].[tbl_Holiday] where Date=@date and CompanyCode=@CompanyCode and PlantID=@PlantID", con);
                        cmd.Parameters.AddWithValue("@date", dates[i11]);
                        cmd.Parameters.AddWithValue("@CompanyCode", CompanyCode);
                        cmd.Parameters.AddWithValue("@PlantID", PlantCode);
                        SqlDataAdapter da = new SqlDataAdapter(cmd);
                        DataTable dtholdiday = new DataTable();
                        da.Fill(dtholdiday);

                        SqlCommand cmd1 = new SqlCommand("SELECT distinct [date],Variantcode,okparts,Linecode,CompanyCode,PlantCode " +
                            "from tbl_daywise_cumulative " +
                            "where Companycode = @company and PlantCode = @plant and Linecode = @Line_code and MachineCode = @machine and [date] = @Date", con);

                        var t = dates[i11].ToString("yyyy-MM-dd");


                        cmd1.Parameters.AddWithValue("@Date", t);
                        cmd1.Parameters.AddWithValue("@Line_code", linecode);
                        cmd1.Parameters.AddWithValue("@company", CompanyCode);
                        cmd1.Parameters.AddWithValue("@plant", PlantCode);
                        cmd1.Parameters.AddWithValue("@machine", machinecode);
                        cmd1.CommandTimeout = 0;
                        SqlDataAdapter da1 = new SqlDataAdapter(cmd1);
                        DataTable dtdaywiseqty = new DataTable();
                        da1.Fill(dtdaywiseqty);

                        ////DAY-WISE production qty

                        ws11.Cell(columnNamadd + "4").Value = dates[i11].ToString("dd-MM-yyyy");
                        ws11.Cell(columnNamadd + "4").Style.Font.Bold = true;
                        ws11.Cell(columnNamadd + "4").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                        ws11.Cell(columnNamadd + "5").Value = machinename;
                        ws11.Cell(columnNamadd + "5").Style.Font.Bold = true;
                        //ws11.Cell(columnNamadd + "5").Style.Alignment.SetVertical(XLAlignmentVerticalValues.Center);

                        var daycheck = dates[i11].ToString("dddd");


                        if (dtholdiday.Rows.Count != 0)
                        {
                            holidayname = dtholdiday.Rows[0][0].ToString();

                            Console.WriteLine(holidayname + "Holiday");

                            ws11.Cell(columnNamadd + "6").Value = holidayname;
                            ws11.Cell(columnNamadd + "6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            ws11.Cell(columnNamadd + "6").Style.Fill.BackgroundColor = XLColor.FromHtml("#f5a742");

                            Console.WriteLine("Holiday details entering completed");
                        }
                        else if (dtdaywiseqty.Rows.Count > 0)
                        {


                            int totalqty = 0;

                            ////--------for cummulative data entering---------////
                            for (int j = 0; j < dtdaywiseqty.Rows.Count; j++)
                            {

                                var aa = dtdaywiseqty.Rows[j][1].ToString();


                                var dd = ws11.CellsUsed(x => x.GetString() == aa);


                                ////to find the respective cell in a n excel
                                //var dd = ws4.Search(aa, CompareOptions.OrdinalIgnoreCase);

                                search(dd, ws11, dtdaywiseqty.Rows[j][2].ToString(), columnNamadd);

                                totalqty += Convert.ToInt32(dtdaywiseqty.Rows[j][2]);

                            }

                            ws11.Cell(columnNamadd + "27").Value = totalqty;

                        }

                        else if (daycheck == "Sunday")
                        {
                            ws11.Cell(columnNamadd + "6").Value = "No Production-Sunday";
                            ws11.Cell(columnNamadd + "6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            ws11.Cell(columnNamadd + "6").Style.Fill.BackgroundColor = XLColor.FromHtml("#f5a742");

                        }

                        else
                        {
                            ws11.Cell(columnNamadd + "6").Value = "No Production";
                            ws11.Cell(columnNamadd + "6").Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                            ws11.Cell(columnNamadd + "6").Style.Fill.BackgroundColor = XLColor.FromHtml("#4efcd1");

                        }



                        int sum111 = sum(columnNamadd);

                        columnNamadd = calculation(sum111 + NoOfMachine);


                        Console.WriteLine("Ended");

                    }



                    IXLCell firstcell = ws11.Cell(columnNamadd1 + "4");

                    IXLCell lastcell = ws11.LastCellUsed();

                    // the range for which you want to add a table style
                    var rngTable1 = ws11.Range(firstcell, lastcell);

                    rngTable1.Style.Border.BottomBorder = XLBorderStyleValues.Thin;

                    rngTable1.Style.Border.LeftBorder = XLBorderStyleValues.Thin;

                    rngTable1.Style.Border.TopBorder = XLBorderStyleValues.Thin;

                    rngTable1.Style.Border.RightBorder = XLBorderStyleValues.Thin;


                    foreach (IXLCell cells in rngTable1.Cells())
                    {
                        if (cells.IsEmpty())
                        {
                            cells.Value = "-";
                            cells.Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);
                        }

                    }


                    ws11.Columns().AdjustToContents();

                    workbook.Save();

                    //------------------------------------------------------
                    //------------------------------------------------------

                    columnName= "H";
                    int sum112 = sum(columnName);

                    string columnNam1 = calculation(sum112 + 1);

                    wss.Cell("B1").Value = columnNam1;

                    ////Date in Index page
                    IXLWorksheet ws12 = workbook.Worksheet(1);
                    string date1 = DateTime.Today.ToString("dd-mm-yyyy");
                    //string date2 = DateTime.Today.AddDays(-1).ToString("dd-MM-yyyy");
                    ws12.Cell("E2").Value = companyName + " - Daily Production Report ";
                    ws12.Cell("J7").Value = machinename;
                    ws12.Cell("E7").Value = date1;
                    ws12.Cell("E8").Value = dat;

                    workbook.Style.Alignment.Horizontal = XLAlignmentHorizontalValues.Center;
                    workbook.Style.Font.Bold = true;

                    //workbook.Worksheet(3).Delete();

                    workbook.Save();
                    //workbook.SaveAs(filepath);


                }
            }
            catch (Exception ex)
            {
                throw ex;
            }
        }

        public static int sum(string columnName)
        {

            int sum = 0;

            for (int i = 0; i < columnName.Length; i++)
            {
                sum *= 26;
                sum += (columnName[i] - 'A' + 1);
            }

            return sum;
        }

        public static String calculation(int a)
        {
            int dividend = a;
            string columnNam = "";
            int modulo;

            while (dividend > 0)
            {
                modulo = (dividend - 1) % 26;
                columnNam = Convert.ToChar(65 + modulo).ToString() + columnNam;
                dividend = (int)((dividend - modulo) / 26);
            }

            return columnNam;

        }

        public static void search(IXLCells dd1, IXLWorksheet sheet, string valuetoenter, string columnNam44)
        {

            try
            {

            

            var a = "";

            foreach (IXLCell cell1 in dd1)
            {
                IXLCell coll = sheet.Column(1).LastCellUsed();

                var s = coll.Address.RowNumber.ToString();

                var df = cell1.Value.ToString();

                if (df == null || df == "")
                {

                    Console.WriteLine("not found so created and added");


                }
                else
                {
                    a = cell1.Address.RowNumber.ToString();

                    // var t = ds3.Tables[i].Rows[j][4].ToString();


                    sheet.Cell(columnNam44 + a).Value = Convert.ToDecimal (valuetoenter);
                    sheet.Cell(columnNam44 + a).Style.Alignment.SetHorizontal(XLAlignmentHorizontalValues.Center);

                }


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