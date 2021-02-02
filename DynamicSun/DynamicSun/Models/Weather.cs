using System;

namespace DynamicSun
{
    public class Weather
    {
        public int Id { get; set; }
        public string ArchiveName { get; set; }
        public string Date { get; set; }
        public string Time { get; set; }
        public string Temp { get; set; }
        public string Wet { get; set; }
        public string DewPoint { get; set; }
        public string Pressure { get; set; }
        public string WindDirect { get; set; }
        public string WindSpeed { get; set; }
        public string CloudCover { get; set; }
        public string LowLimitCloud { get; set; }
        public string HorizontalVisibility { get; set; }
        public string WeatherEffect { get; set; }
    }

}