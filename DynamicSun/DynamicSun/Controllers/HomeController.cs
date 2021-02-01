using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Web;
using System.Web.Mvc;

namespace DynamicSun.Controllers
{

    public class HomeController : Controller
    {
        WeatherContext db = new WeatherContext();
        public ActionResult Index()
        {
            ViewBag.Archives = db.Archives;
            return View();
        }

        public ActionResult About()
        {
            ViewBag.Message = "Your application description page.";

            return View();
        }



        public ActionResult LoadFile(IEnumerable<HttpPostedFileBase> fileUpload)
        {
            if (fileUpload != null)
            {
                foreach (var file in fileUpload)
                {
                    if (file == null) continue;
                    string path = AppDomain.CurrentDomain.BaseDirectory + "App_Data\\";
                    string filename = Path.GetFileName(file.FileName);
                    if (filename != null) file.SaveAs(Path.Combine(path, filename));
                    Thread ReadFile = new Thread(new ParameterizedThreadStart(LoadFileInBd));
                    ReadFile.Start(path + filename);

                }
            }
            return View();
        }

        public static void LoadFileInBd(object obj)
        {
            string path = (string)obj;
            XSSFWorkbook hssfwb;
            using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
            {
                hssfwb = new XSSFWorkbook(file);

            }
            var a = hssfwb.NumberOfSheets;
            for (int i=0; i<=a; i++)
            {
                ISheet sheet = hssfwb.GetSheetAt(i);
                for (int row = 5; row <= sheet.LastRowNum; row++)
                {
                    if (sheet.GetRow(row) != null)
                    {
                        var ad = sheet.GetRow(row);
                    }
                }
            }
        }
    }
}