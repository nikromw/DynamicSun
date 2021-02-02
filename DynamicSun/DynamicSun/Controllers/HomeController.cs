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
       static WeatherContext db = new WeatherContext();
        public ActionResult Index()
        {
            LoadArchives();
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
                    LoadArchives();
                }
            }
            return View();
        }

        public void LoadArchives()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string[] files = Directory.GetFiles(path + "\\App_Data\\");
            List<string> archives = new List<string>();
            foreach (var file in files)
            {
                archives.Add(file.Split('\\').Last());
            }
            ViewBag.Archives = archives;
        }

        public static void LoadFileInBd(object obj)
        {
            string path = (string)obj;
            XSSFWorkbook hssfwb;
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);

                }
                var a = hssfwb.NumberOfSheets;
                for (int i = 0; i < a; i++)
                {
                    ISheet sheet = hssfwb.GetSheetAt(i);
                    for (int row = 5; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            try
                            {
                                Weather weather = new Weather();
                                var RowElements = sheet.GetRow(row).Cells;
                                weather.Date = RowElements[0].ToString();
                                weather.Time = RowElements[1].ToString();
                                weather.Temp = RowElements[2].ToString();
                                weather.Wet = RowElements[3].ToString();
                                weather.DewPoint = RowElements[4].ToString();
                                weather.Pressure = RowElements[5].ToString();
                                weather.WindDirect = RowElements[6].ToString();
                                weather.WindSpeed = RowElements[7].ToString();
                                weather.CloudCover = RowElements[8].ToString();
                                weather.LowLimitCloud = RowElements[9].ToString();
                                weather.ArchiveName = path.Split('\\').Last();
                                weather.HorizontalVisibility = RowElements[10].ToString();
                                weather.WeatherEffect = RowElements.ElementAtOrDefault(11) == null ? "" : RowElements[11].ToString();
                                db.Weathers.Add(weather);
                            }
                            catch (Exception e)
                            { }
                        }
                    }
                }
                db.SaveChanges();
            }
            catch (Exception e)
            { }
        }
    }
}