﻿using NPOI.HSSF.UserModel;
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
        static WeatherContext db = new WeatherContext();
        static List<Weather> weatherFromDb = new List<Weather>();
        static Dictionary<string, bool> archives = new Dictionary<string, bool>();
        static string lastSortYear = "";
        static string lastSortMonth = "";
        public ActionResult Index(string archive)
        {
            ViewBag.Amount = 10;
            LoadArchiveFromDb(archive);
            LoadArchives();
            return View();
        }

        public ActionResult ViewWeather(string YearSort, string MonthSort, int page = 1)
        {
            int pageSize = 15;
            var sortList = weatherFromDb;
            try
            {
                if ((YearSort != null && MonthSort != null && YearSort != "" && MonthSort != "") || (lastSortYear != "" && lastSortMonth != "" && MonthSort != ""))
                {
                    if ((YearSort != null && MonthSort != null && YearSort != "" && MonthSort != ""))
                    {
                        lastSortYear = YearSort;
                        lastSortMonth = MonthSort;
                    }
                    sortList = weatherFromDb.Where(i => (i.Date.Split('.')[2] == YearSort
                                                && i.Date.Split('.')[1] == MonthSort)
                                                || (i.Date.Split('.')[1] == lastSortMonth
                                                && i.Date.Split('.')[2] == lastSortYear)).ToList();
                }
                else if ((YearSort != null && YearSort != "") || lastSortYear != "")
                {
                    if (YearSort != "" && YearSort != null) lastSortYear = YearSort;
                    sortList = weatherFromDb.Where(i => i.Date.Split('.')[2] == YearSort || i.Date.Split('.')[2] == lastSortYear).ToList();

                }
            }
            catch (Exception e)
            { }
            IEnumerable<Weather> WeatherPages = sortList.Skip((page - 1) * pageSize).Take(pageSize);
            PageInfo pageInfo = new PageInfo { PageNumber = page, PageSize = pageSize, TotalItems = sortList.Count };
            IndexViewModel ivm = new IndexViewModel { PageInfo = pageInfo, Weathers = WeatherPages };
            return View(ivm);
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
                    if (filename != null)
                    {
                        file.SaveAs(Path.Combine(path, filename));
                        LoadFileInBd(path + filename);
                    }
                    LoadArchives();
                }
            }
            return View();
        }

        public void LoadArchives()
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string[] files = Directory.GetFiles(path + "\\App_Data\\");
            foreach (var file in files)
            {
                if (!archives.ContainsKey(file.Split('\\').Last()))
                {
                    archives.Add(file.Split('\\').Last(), false);
                }
            }
            ViewBag.Archives = archives;
        }

        public void LoadFileInBd(string path)
        {
            XSSFWorkbook hssfwb;
            try
            {
                using (FileStream file = new FileStream(path, FileMode.Open, FileAccess.Read))
                {
                    hssfwb = new XSSFWorkbook(file);
                }
                var a = hssfwb.NumberOfSheets;
                var weatherFromDb = db.Weathers.ToList();
                for (int i = 0; i < a; i++)
                {
                    ISheet sheet = hssfwb.GetSheetAt(i);
                    for (int row = 5; row <= sheet.LastRowNum; row++)
                    {
                        if (sheet.GetRow(row) != null)
                        {
                            try
                            {
                                Weather weather;
                                var RowElements = sheet.GetRow(row).Cells;
                                if (weatherFromDb.Exists(W => W.Date == RowElements[0].ToString()))
                                {
                                    weather = weatherFromDb.Where(W => W.Date == RowElements[0].ToString()).FirstOrDefault();
                                }
                                else
                                {
                                    weather = new Weather();
                                }
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
                    db.SaveChanges();
                }
            }
            catch (Exception e)
            { }
        }
        public void LoadArchiveFromDb(string archive)
        {
            try
            {
                weatherFromDb.RemoveAll(i => i.ArchiveName == archive);
                weatherFromDb.AddRange(db.Weathers.Where(i => i.ArchiveName == archive).ToList());
                ViewBag.weatherFromdb = weatherFromDb;
            }
            catch (Exception e)
            { }
        }


    }
}