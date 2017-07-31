using System;
using System.Collections.Generic;
using System.Linq;
using System.Web;
using System.Web.Mvc;
using ExcelGeneratorOpenXML;
using System.IO;

namespace WebMVCExcel.Controllers
{
    public class HomeController : Controller
    {
        public ActionResult Index()
        {
            return View();
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
        public ActionResult GetExcelFile()
        {
            MemoryStream excelStream = new MemoryStream();
            var excelFile = new ExcelService();
            bool fileCreated = excelFile.CreateSpreadSheet();
            excelStream = excelFile.SpreadsheetStream;

            HttpContext.Response.Clear();
            HttpContext.Response.AddHeader("content-disposition", "attachment;filename=test.xls");
            excelStream.WriteTo(HttpContext.Response.OutputStream);
            excelStream.Close();
            HttpContext.Response.End();
            
            return View();
        }
    }
}