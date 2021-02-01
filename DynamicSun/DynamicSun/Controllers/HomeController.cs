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
                        Weather weather = new Weather();
                        var RowElements = sheet.GetRow(row).Cells;
                        weather.Date = new DateTime(Convert.ToInt32(RowElements[0].ToString().Split('.')[2]),
                                                    Convert.ToInt32(RowElements[0].ToString().Split('.')[1]),
                                                    Convert.ToInt32(RowElements[0].ToString().Split('.')[0]),
                                                    Convert.ToInt32(RowElements[1].ToString().Split(':')[0]),
                                                    Convert.ToInt32(RowElements[1].ToString().Split(':')[1]),
                                                    0);
                        weather.Temp = Convert.ToDouble(RowElements[2].ToString());
                        weather.Wet = Convert.ToDouble(RowElements[3].ToString());
                        weather.DewPoint = Convert.ToDouble(RowElements[4].ToString());
                        weather.Pressure = Convert.ToInt32(RowElements[5].ToString());
                        weather.WindDirect = RowElements[6].ToString();
                        weather.WindSpeed = Convert.ToDouble(RowElements[7].ToString());
                        weather.CloudCover = Convert.ToDouble(RowElements[8].ToString());
                        weather.LowLimitCloud = Convert.ToDouble(RowElements[9].ToString());
                        weather.HorizontalVisibility = Convert.ToDouble(RowElements[10].NumericCellValue == null ? " ": RowElements[10].StringCellValue);
                        weather.WeatherEffect = RowElements[11].ToString();
                    }
                }
            }
        }
    }
}