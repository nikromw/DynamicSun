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
        public ActionResult Index(string archive)
        {
            if (Session["Archives"] == null && Session["weatherFromdb"] == null)
            {
                Session["Archives"] = new Dictionary<string, bool>();
                Session["weatherFromdb"] = new List<Weather>();
            }
            using (WeatherContext db = new WeatherContext())
            {
                foreach (var Ar in db.Archives.ToList())
                {
                    if (!((Dictionary<string, bool>)Session["Archives"]).ContainsKey(Ar.Name))
                    {
                        ((Dictionary<string, bool>)Session["Archives"]).Add(Ar.Name, false);
                    }
                }
            }
            ViewBag.Amount = 10;
            LoadArchiveFromDb(archive);
            return View();
        }

        public ActionResult ViewWeather(int YearFilt = 0, int MonthFiltr = 0, int page = 1)
        {
            int pageSize = 15;
            if (Session["lastFiltrYear"]== null || Session["lastFiltrMonth"]==null )
            {
                Session["lastFiltrMonth"] = 0;
                Session["lastFiltrYear"] = 0;
            }
            List<Weather> sortList = (List<Weather>)Session["weatherFromdb"] == null ? new List<Weather>() : (List<Weather>)Session["weatherFromdb"];
            if (YearFilt != 0 && MonthFiltr != 0 || (int)Session["lastFiltrYear"] != 0 && (int)Session["lastFiltrMonth"] != 0)
            {
                if (YearFilt != 0 && MonthFiltr != 0)
                {
                    Session["lastFiltrMonth"] = Convert.ToInt32(MonthFiltr);
                    Session["lastFiltrYear"] = Convert.ToInt32(YearFilt);
                }

                sortList = sortList.Where(i => (i.Date.Value.Year == YearFilt && i.Date.Value.Month == MonthFiltr)
                || (i.Date.Value.Year == (int)Session["lastFiltrYear"] && i.Date.Value.Month == (int)Session["lastFiltrMonth"])).ToList();
            }
            else if (YearFilt != 0 || (int)Session["lastFiltrYear"] != 0)
            {
                if (YearFilt != 0)
                {
                    Session["lastFiltrYear"] = Convert.ToInt32(YearFilt);
                }
                sortList = sortList.Where(i => i.Date.Value.Year == YearFilt || i.Date.Value.Year == (int)Session["lastFiltrYear"]).ToList();
            }
            sortList.OrderBy(i => i.Date);
            ViewBag.FilteredWeather = sortList.OrderBy(i => i.Date);
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
                    if (Session["Archives"]!= null && !((Dictionary<string, bool>)Session["Archives"]).ContainsKey(file.FileName))
                    {
                        LoadFileInBd(file);
                    }
                }
            }
            return View();
        }

        public void LoadFileInBd(object obj)
        {
            HttpPostedFileWrapper path = (HttpPostedFileWrapper)obj;
            XSSFWorkbook hssfwb;


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
                        {
                            using (FileStream fstream = new FileStream($"Log.txt", FileMode.OpenOrCreate))
                            {
                                byte[] array = System.Text.Encoding.Default.GetBytes($"Ошибка чтения файла: {path.FileName}, страница : {i} , строка {row}  ") ; 
                                fstream.Write(array, 0, array.Length);
                            }
                        }
                    }
                    db.SaveChanges();
            }
        }
    }
    public void LoadArchiveFromDb(string archive)
    {
        if (archive != null && Session["Archives"]!= null && !((Dictionary<string, bool>)Session["Archives"])[archive])
        {
            using (WeatherContext db = new WeatherContext())
            {
                ((Dictionary<string, bool>)Session["Archives"])[archive] = true;
                ((List<Weather>)(Session["weatherFromdb"])).AddRange(db.Weathers.Where(i => i.ArchiveName == archive).ToList());
            }
        }
    }
}
}