using ExcelFileDemo.Models;
using ExcelFileDemo.ViewModel;
using LinqToExcel;
using Microsoft.AspNet.Identity.EntityFramework;
using Spire.Xls;
using System;
using System.Collections.Generic;
using System.Data;
using System.Data.Common;
using System.Data.Entity;
using System.Data.OleDb;
using System.Data.SqlClient;
using System.IO;
using System.Linq;
using System.Web;
using System.Web.Mvc;

namespace ExcelFileDemo.Controllers
{
    public class HomeController : Controller
    {
        private readonly ApplicationDbContext _context;
        public HomeController()
        {
            _context = new ApplicationDbContext();
        }
        public ActionResult Index()
        {
            var result = _context.Database.SqlQuery<string>("SELECT TABLE_NAME FROM INFORMATION_SCHEMA.TABLES WHERE TABLE_TYPE = 'BASE TABLE' AND TABLE_CATALOG = 'ExcelFileDemoDb'").ToList();

            var vm = new HomeViewModel
            {
                Tables = result
            };
            return View(vm);
        }
        public JsonResult GetTableData(string tableName)
        {
            var result = _context.Database.SqlQuery<TableData>("SELECT COLUMN_NAME,DATA_TYPE FROM INFORMATION_SCHEMA.COLUMNS WHERE TABLE_NAME = '" + tableName + "'").ToList();

            return Json(result, JsonRequestBehavior.AllowGet);
        }
        public JsonResult GetTableExcelFileData(string filePath)
        {
            string file = @"C:\Users\user\Downloads\Districts النشط - Copy.xlsx";
            var path = Path.GetFullPath(filePath).Replace(@"\\", @"\");
            var newBook = new Workbook();
            newBook.LoadFromFile(file);
            var sheet = newBook.Worksheets[0];
            var sheetCount = sheet.Columns.Count();

            return null;
            //return Json(result, JsonRequestBehavior.AllowGet);
        }
        [HttpPost]
        public JsonResult UploadFile(HttpPostedFileBase file)
        {
            OleDbConnection connection = new OleDbConnection();
            try
            {
                if (file.ContentLength > 0)
                {
                    string _FileName = Path.GetFileName(file.FileName);
                    string _path = Path.Combine(Server.MapPath("~/Content/UploadedFiles"), _FileName);
                    file.SaveAs(_path);

                    //var newBook = new Workbook();
                    //newBook.LoadFromFile(_path);
                    //var sheet = newBook.Worksheets[0];


                    var excelConnectionString = "";
                    if (file.FileName.EndsWith(".xls"))
                    {
                        excelConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", _path);
                    }
                    else if (file.FileName.EndsWith(".xlsx"))
                    {
                        excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", _path);
                    }

                    //string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties =Excel 8.0", _path);

                    connection.ConnectionString = excelConnectionString;
                    OleDbCommand command = new OleDbCommand("select * from[Sheet1$]", connection);
                    connection.Open();
                    DbDataReader dr = command.ExecuteReader();
                    var columns = new List<string>();

                    for (int i = 0; i < dr.FieldCount; i++)
                    {
                        columns.Add(dr.GetName(i));
                    }
                    dr.GetSchemaTable();
                    string sqlConnectionString = @"Data Source =(LocalDb)\MSSQLLocalDB;Initial Catalog=ExcelFileDemoDb;Integrated Security=True";
                    SqlBulkCopy bulkInsert = new SqlBulkCopy(sqlConnectionString);
                    bulkInsert.ColumnMappings.Add("Id", "Id");
                    //bulkInsert.ColumnMappings.Add("City", "IsDeleted");
                    bulkInsert.ColumnMappings.Add("City", "Name");
                    bulkInsert.ColumnMappings.Add("deleted", "IsDeleted");
                    bulkInsert.DestinationTableName = "Cities";
                    bulkInsert.WriteToServer(dr);


                    //var adapter = new OleDbDataAdapter("SELECT * FROM [Sheet1$]", connectionString);
                    //var ds = new DataSet();

                    //adapter.Fill(ds, "ExcelTable");

                    //DataTable dtable = ds.Tables["ExcelTable"];

                    //string sheetName = "Sheet1";

                    //var excelFile = new ExcelQueryFactory(_path);


                    //var artistAlbums = from a in excelFile.Worksheet<User>(sheetName) select a;
                }
                return Json("success", JsonRequestBehavior.AllowGet);
            }
            catch (Exception ex)
            {
                ViewBag.Message = "File upload failed!!";
                return Json("fail", JsonRequestBehavior.AllowGet);
            }
            finally
            {
                connection.Close();
            }
        }
        [HttpPost]
        public JsonResult UploadExcelFile()
        {
            if (Request.Files.Count > 0)
            {
                try
                {
                    //  Get all files from Request object  
                    HttpFileCollectionBase files = Request.Files;
                    //for (int i = 0; i < files.Count; i++)
                    //{
                    //string path = AppDomain.CurrentDomain.BaseDirectory + "Uploads/";  
                    //string filename = Path.GetFileName(Request.Files[i].FileName);  

                    HttpPostedFileBase file = files[0];
                    string fname;

                    // Checking for Internet Explorer  
                    if (Request.Browser.Browser.ToUpper() == "IE" || Request.Browser.Browser.ToUpper() == "INTERNETEXPLORER")
                    {
                        string[] testfiles = file.FileName.Split(new char[] { '\\' });
                        fname = testfiles[testfiles.Length - 1];
                    }
                    else
                    {
                        fname = file.FileName;
                    }

                    // Get the complete folder path and store the file inside it.  
                    fname = Path.Combine(Server.MapPath("~/Content/UploadedFiles"), fname);
                    file.SaveAs(fname);

                    var excelColumns = ReadExcelColumn(fname);
                    //}
                    // Returns message that successfully uploaded  
                    return Json(excelColumns, JsonRequestBehavior.AllowGet);
                }
                catch (Exception ex)
                {
                    return Json("Error occurred. Error details: " + ex.Message);
                }
            }
            else
            {
                return Json("No files selected.");
            }
        }
        List<string> ReadExcelColumn(string path)
        {
            OleDbConnection connection = new OleDbConnection();
            try
            {
                string fileName = Path.GetFileName(path);
                var excelConnectionString = "";
                if (fileName.EndsWith(".xls"))
                {
                    excelConnectionString = string.Format("Provider=Microsoft.Jet.OLEDB.4.0; data source={0}; Extended Properties=Excel 8.0;", path);
                }
                else if (fileName.EndsWith(".xlsx"))
                {
                    excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties=\"Excel 12.0 Xml;HDR=YES;IMEX=1\";", path);
                }
                excelConnectionString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + path + ";Extended Properties=Excel 12.0;";
                //string excelConnectionString = string.Format("Provider=Microsoft.ACE.OLEDB.12.0;Data Source={0};Extended Properties =Excel 8.0", _path);

                connection.ConnectionString = excelConnectionString;
                OleDbCommand command = new OleDbCommand("select * from[Sheet1$]", connection);
                connection.Open();
                DbDataReader dr = command.ExecuteReader();
                var columns = new List<string>();

                for (int i = 0; i < dr.FieldCount; i++)
                {
                    columns.Add(dr.GetName(i));
                }
                return columns;
            }
            catch (Exception ex)
            {


            }
            finally
            {
                connection.Close();
            }
            return null;

        }
        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }

        public ActionResult Contact()
        {
            ViewBag.Message = "Your contact page.";

            return View();
        }
    }
}