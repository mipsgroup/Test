using System;
using System.Collections.Generic;
using System.Diagnostics;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.Extensions.Logging;
using MipsTester.Models;
using Microsoft.AspNetCore.Hosting;
using Microsoft.AspNetCore.Http;
using System.IO;
using ExcelDataReader;
using System.Data;
using System.Data.SqlClient;
using Microsoft.Office.Interop.Excel;
using DataTable = System.Data.DataTable;

namespace MipsTester.Controllers
{
    public class HomeController : Controller
    {
        private readonly ILogger<HomeController> _logger;
        private readonly IWebHostEnvironment _env;

        public HomeController(ILogger<HomeController> logger, IWebHostEnvironment env)
        {
            _logger = logger;
            _env = env;
        }
        public IActionResult Index()
        {
            return View();
        }

        public IActionResult Privacy()
        {
            return View();
        }

        [ResponseCache(Duration = 0, Location = ResponseCacheLocation.None, NoStore = true)]
        public IActionResult Error()
        {
            //new change
            return View(new ErrorViewModel { RequestId = Activity.Current?.Id ?? HttpContext.TraceIdentifier });
        }

        public ActionResult UploadFile()
        {
            return View();
        }

        [HttpPost]
        [RequestFormLimits(MultipartBodyLengthLimit = 209715200)]
        [RequestSizeLimit(209715200)]
        public ActionResult Index(FileUpload fileUpload)
        {
            if (fileUpload.FormFiles != null)
            {
                List<DataTable> dt = new List<DataTable>();
                foreach (IFormFile formFile in fileUpload.FormFiles)
                {
                    string filePath = $"{_env.WebRootPath}/files/{formFile.FileName}";
                    using (var stream = System.IO.File.Create(filePath))
                    {
                        formFile.CopyTo(stream);
                    }

                    dt.Add(GetDataTableFromExcel(filePath,true));
                }

                //compare 2 tables
                ViewBag.Data = getDifferentRecords(dt[0], dt[1]);
            }
            //return Redirect("/");
            return View();
        }

        public static DataTable CompareTwoDataTable(DataTable dt1, DataTable dt2)
        {

            dt1.Merge(dt2);

            DataTable d3 = dt2.GetChanges();


            return d3;

        }
        public DataTable getDifferentRecords(DataTable FirstDataTable, DataTable SecondDataTable)
        {
            //Create Empty Table   
            DataTable ResultDataTable = new DataTable("ResultDataTable");

            //use a Dataset to make use of a DataRelation object   
            using (DataSet ds = new DataSet())
            {
                //Add tables   
                ds.Tables.AddRange(new DataTable[] { FirstDataTable.Copy(), SecondDataTable.Copy() });

                //Get Columns for DataRelation   
                DataColumn[] firstColumns = new DataColumn[ds.Tables[0].Columns.Count];
                for (int i = 0; i < firstColumns.Length; i++)
                {
                    firstColumns[i] = ds.Tables[0].Columns[i];
                }

                DataColumn[] secondColumns = new DataColumn[ds.Tables[1].Columns.Count];
                for (int i = 0; i < secondColumns.Length; i++)
                {
                    secondColumns[i] = ds.Tables[1].Columns[i];
                }

                //Create DataRelation   
                DataRelation r1 = new DataRelation(string.Empty, firstColumns, secondColumns, false);
                ds.Relations.Add(r1);

                DataRelation r2 = new DataRelation(string.Empty, secondColumns, firstColumns, false);
                ds.Relations.Add(r2);

                //Create columns for return table   
                for (int i = 0; i < FirstDataTable.Columns.Count; i++)
                {
                    ResultDataTable.Columns.Add(FirstDataTable.Columns[i].ColumnName, FirstDataTable.Columns[i].DataType);
                }

                //If FirstDataTable Row not in SecondDataTable, Add to ResultDataTable.   
                ResultDataTable.BeginLoadData();
                foreach (DataRow parentrow in ds.Tables[0].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r1);
                    if (childrows == null || childrows.Length == 0)
                        ResultDataTable.LoadDataRow(parentrow.ItemArray, true);
                }

                //If SecondDataTable Row not in FirstDataTable, Add to ResultDataTable.   
                foreach (DataRow parentrow in ds.Tables[1].Rows)
                {
                    DataRow[] childrows = parentrow.GetChildRows(r2);
                    if (childrows == null || childrows.Length == 0)
                        ResultDataTable.LoadDataRow(parentrow.ItemArray, true);
                }
                ResultDataTable.EndLoadData();
            }

            return ResultDataTable;
        }


        public DataTable READExcel(string path)
        {
            Microsoft.Office.Interop.Excel.Application objXL = null;
            Microsoft.Office.Interop.Excel.Workbook objWB = null;
            objXL = new Microsoft.Office.Interop.Excel.Application();
            objWB = objXL.Workbooks.Open(path);
            var objSHT = (Worksheet)objWB.Worksheets[1];

            int rows = objSHT.UsedRange.Rows.Count;
            int cols = objSHT.UsedRange.Columns.Count;
            var dt = new DataTable();
            int noofrow = 1;

            for (int c = 1; c <= cols; c++)
            {
                string colname = objSHT.Cells[1, c].ToString();
                dt.Columns.Add(colname);
                noofrow = 2;
            }

            for (int r = noofrow; r <= rows; r++)
            {
                DataRow dr = dt.NewRow();
                for (int c = 1; c <= cols; c++)
                {
                    dr[c - 1] = objSHT.Cells[r, c].ToString();
                }

                dt.Rows.Add(dr);
            }

            objWB.Close();
            objXL.Quit();
            return dt;
        }

        public void Do(string path1)
        {
            string query = null;
            string connString = "";

            var extension = System.IO.Path.GetExtension(path1).ToLower();


            string[] validFileTypes = { ".xls", ".xlsx", ".csv" };


            if (validFileTypes.Contains(extension))
            {

                if (extension == ".csv")
                {
                    DataTable dt = Utility.ConvertCSVtoDataTable(path1);
                    ViewBag.Data = dt;
                }
                //Connection String to Excel Workbook  
                else if (extension.Trim() == ".xls")
                {
                    connString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" + path1 + ";Extended Properties=\"Excel 8.0;HDR=Yes;IMEX=2\"";
                    DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                    ViewBag.Data = dt;
                }
                else if (extension.Trim() == ".xlsx")
                {
                    connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path1 + ";Extended Properties=\"Excel 12.0;HDR=Yes;IMEX=2\"";
                    DataTable dt = Utility.ConvertXSLXtoDataTable(path1, connString);
                    ViewBag.Data = dt;
                }

            }
        }

        private static DataTable GetDataTableFromExcel(string path, bool hasHeader = true)
        {
            using (var pck = new OfficeOpenXml.ExcelPackage())
            {
                using (var stream = System.IO.File.OpenRead(path))
                {
                    pck.Load(stream);
                }
                var ws = pck.Workbook.Worksheets[0]; // Second .First();
                DataTable tbl = new DataTable();
                foreach (var firstRowCell in ws.Cells[1, 1, 1, ws.Dimension.End.Column])
                {
                    tbl.Columns.Add(hasHeader ? firstRowCell.Text : string.Format("Column {0}", firstRowCell.Start.Column));
                }
                var startRow = hasHeader ? 2 : 1;
                for (int rowNum = startRow; rowNum <= ws.Dimension.End.Row; rowNum++)
                {
                    var wsRow = ws.Cells[rowNum, 1, rowNum, ws.Dimension.End.Column];
                    DataRow row = tbl.Rows.Add();
                    foreach (var cell in wsRow)
                    {
                        row[cell.Start.Column -1] = cell.Text;
                    }
                }
                return tbl;
            }
        }
        //private IActionResult ReadExcelFileAsync(IFormFile file)
        //{
        //    if (file == null || file.Length == 0)
        //        return Content("File Not Selected");

        //    string fileExtension = Path.GetExtension(file.FileName);

        //    if (fileExtension == ".xls" || fileExtension == ".xlsx")
        //    {
        //        var rootFolder = @"D:\Files";
        //        var fileName = file.FileName;
        //        var filePath = Path.Combine(rootFolder, fileName);
        //        var fileLocation = new FileInfo(filePath);




        //        using (var stream = file.OpenReadStream()) // File.Open("Book.xlsx", FileMode.Open, FileAccess.Read))
        //        {
        //            using (var reader = ExcelReaderFactory.CreateReader(stream))
        //            {
        //                do
        //                {
        //                    DataTable dt = new DataTable();

        //                    while (reader.Read()) //Each ROW
        //                    {
        //                        for (int column = 0; column < reader.FieldCount; column++)
        //                        {
        //                            //Console.WriteLine(reader.GetString(column));//Will blow up if the value is decimal etc. 
        //                            Console.WriteLine(reader.GetValue(column));//Get Value returns object
        //                        }
        //                    }
        //                } while (reader.NextResult()); //Move to NEXT SHEET

        //            }
        //        }





        //        //    using (var fileStream = new FileStream(filePath, FileMode.Create))
        //        //    {
        //        //        await file.CopyToAsync(fileStream);
        //        //    }

        //        //    if (file.Length <= 0)
        //        //        return BadRequest(GlobalValidationMessage.FileNotFound);

        //        //    using (ExcelPackage package = new ExcelPackage(fileLocation))
        //        //    {
        //        //        ExcelWorksheet workSheet = package.Workbook.Worksheets["Table1"];
        //        //        //var workSheet = package.Workbook.Worksheets.First();
        //        //        int totalRows = workSheet.Dimension.Rows;

        //        //        var DataList = new List<Customers>();

        //        //        for (int i = 2; i <= totalRows; i++)
        //        //        {
        //        //            DataList.Add(new Customers
        //        //            {
        //        //                CustomerName = workSheet.Cells[i, 1].Value.ToString(),
        //        //                CustomerEmail = workSheet.Cells[i, 2].Value.ToString(),
        //        //                CustomerCountry = workSheet.Cells[i, 3].Value.ToString()
        //        //            });
        //        //        }

        //        //        _db.Customers.AddRange(customerList);
        //        //        _db.SaveChanges();
        //        //    }
        //        //}

        //        return Ok();
        //    }

        //}

        public static class Utility
        {
            public static DataTable ConvertCSVtoDataTable(string strFilePath)
            {
                DataTable dt = new DataTable();
                using (StreamReader sr = new StreamReader(strFilePath))
                {
                    string[] headers = sr.ReadLine().Split(',');
                    foreach (string header in headers)
                    {
                        dt.Columns.Add(header);
                    }

                    while (!sr.EndOfStream)
                    {
                        string[] rows = sr.ReadLine().Split(',');
                        if (rows.Length > 1)
                        {
                            DataRow dr = dt.NewRow();
                            for (int i = 0; i < headers.Length; i++)
                            {
                                dr[i] = rows[i].Trim();
                            }
                            dt.Rows.Add(dr);
                        }
                    }

                }


                return dt;
            }


            public static DataTable ConvertXSLXtoDataTable(string strFilePath, string connString)
            {
                var Conn = new SqlConnection(connString);
                DataTable dt = new DataTable();
                try
                {

                    Conn.Open();
                    using (var cmd = new SqlCommand("SELECT * FROM [Sheet1$]", Conn))
                    {
                        var da = new SqlDataAdapter();
                        da.SelectCommand = cmd;
                        DataSet ds = new DataSet();
                        da.Fill(ds);

                        dt = ds.Tables[0];
                    }
                }
                catch
                {
                }
                finally
                {

                    Conn.Close();
                }

                return dt;

            }
        }
    }
}
