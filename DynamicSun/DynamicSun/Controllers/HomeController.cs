using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using PagedList;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;
using System.Web.UI;

namespace DynamicSun.Controllers
{

    public class HomeController : Controller
    {
        static List<Weather> weatherFromDb = new List<Weather>();
        static Dictionary<string, bool> archives = new Dictionary<string, bool>();
        static int lastFiltrYear = 0;
        static int lastFiltrMonth = 0;
        public ActionResult Index(string archive)
        {
            using (WeatherContext db = new WeatherContext())
            {
                foreach (var Ar in db.Archives.ToList())
                {
                    if (!archives.ContainsKey(Ar.Name))
                        archives.Add(Ar.Name, false);
                }
            }
            ViewBag.Amount = 10;
            LoadArchiveFromDb(archive);
            ViewBag.Archives = archives;
            return View();
        }

        public ActionResult ViewWeather(int YearFilt = 0, int MonthFiltr = 0, int page = 1)
        {
            int pageSize = 15;
            var sortList = weatherFromDb;
            if (YearFilt != 0 && MonthFiltr != 0 || lastFiltrYear != 0 && lastFiltrMonth != 0)
            {
                if (YearFilt != 0 && MonthFiltr != 0)
                {
                    lastFiltrMonth = Convert.ToInt32(MonthFiltr);
                    lastFiltrYear = Convert.ToInt32(YearFilt);
                }

                sortList = sortList.Where(i => (i.Date.Value.Year == YearFilt && i.Date.Value.Month == MonthFiltr)
                || (i.Date.Value.Year == lastFiltrYear && i.Date.Value.Month == lastFiltrMonth)).ToList();
            }
            else if (YearFilt != 0 || lastFiltrYear != 0)
            {
                if (YearFilt != 0)
                {
                    lastFiltrYear = Convert.ToInt32(YearFilt);
                }
                sortList = sortList.Where(i => i.Date.Value.Year == YearFilt || i.Date.Value.Year == lastFiltrYear).ToList();
            }
            ViewBag.YearFiltr = lastFiltrYear;
            ViewBag.MonthFiltr = lastFiltrMonth;
            IEnumerable<Weather> WeatherPages = sortList.Skip((page - 1) * pageSize).Take(pageSize);
            PageInfo pageInfo = new PageInfo { PageNumber = page, PageSize = pageSize, TotalItems = sortList.Count };
            IndexViewModel ivm = new IndexViewModel { PageInfo = pageInfo, Weathers = WeatherPages };
            return View(ivm);
        }



        public async Task<ActionResult> LoadFile(IEnumerable<HttpPostedFileBase> fileUpload)
        {
            if (fileUpload != null)
            {
                foreach (var file in fileUpload)
                {
                    if (file == null) continue;
                    string path = AppDomain.CurrentDomain.BaseDirectory + "App_Data\\";
                    string filename = Path.GetFileName(file.FileName);
                    if (filename != null)
                    {
                        LoadFileInBd(file);
                    }
                    ViewBag.Archives = archives;
                }
            }
            return View();
        }

        public void LoadFileInBd(object obj)
        {
            HttpPostedFileWrapper path = (HttpPostedFileWrapper)obj;
            XSSFWorkbook hssfwb;
            try
            {
               
                hssfwb = new XSSFWorkbook(path.InputStream);
                var a = hssfwb.NumberOfSheets;

                Archive archive = new Archive();
                archive.Name = path.FileName;
                using (WeatherContext db = new WeatherContext())
                {
                    db.Archives.Add(archive);
                    var weatherFromDb = db.Weathers.ToList();
                    for (int i = 0; i < a; i++)
                    {
                        ISheet sheet = hssfwb.GetSheetAt(i);
                        for (int row = 5; row <= sheet.LastRowNum; row++)
                        {
                            try
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
                                    weather.Temp = Convert.ToDouble(RowElements[2].CellType == CellType.String ? RowElements[2].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[2].NumericCellValue.ToString());
                                    weather.Wet = Convert.ToDouble(RowElements[3].CellType == CellType.String ? RowElements[3].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[3].NumericCellValue.ToString());
                                    weather.DewPoint = Convert.ToDouble(RowElements[4].CellType == CellType.String ? RowElements[4].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[4].NumericCellValue.ToString());
                                    weather.Pressure = Convert.ToDouble(RowElements[5].CellType == CellType.String ? RowElements[5].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[5].NumericCellValue.ToString());
                                    weather.WindDirect = RowElements[6].ToString();
                                    weather.WindSpeed = Convert.ToDouble(RowElements[7].CellType == CellType.String ? RowElements[7].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[7].NumericCellValue.ToString());
                                    weather.CloudCover = Convert.ToDouble(RowElements[8].CellType == CellType.String ? RowElements[8].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[8].NumericCellValue.ToString());
                                    weather.LowLimitCloud = Convert.ToDouble(RowElements[9].CellType == CellType.String ? RowElements[9].RichStringCellValue.String == " " ? null : RowElements[7].RichStringCellValue.String : RowElements[9].NumericCellValue.ToString());
                                    weather.HorizontalVisibility = RowElements[10].ToString();
                                    weather.WeatherEffect = RowElements.ElementAtOrDefault(11) == null ? "" : RowElements[11].StringCellValue;
                                    weather.ArchiveName = path.FileName;
                                    db.Weathers.Add(weather);
                                }
                            }
                            catch (Exception e)
                            { }
                        }
                        db.SaveChanges();
                    }
                }
            }
            catch (Exception e)
            { }
        }

        //Выгружаю только те архивы, которые нужны. И запретил повторную выгрузку.
        //Вся работа с погодой будет только со статическим массивом, без постоянной выгрузки из бд
        public void LoadArchiveFromDb(string archive)
        {
            if (archive != null && !archives[archive])
            {
                using (WeatherContext db = new WeatherContext())
                {
                    archives[archive] = true;
                    weatherFromDb.AddRange(db.Weathers.Where(i => i.ArchiveName == archive).ToList());
                    ViewBag.weatherFromdb = weatherFromDb;
                }
            }
        }
    }
}