using System;

namespace DynamicSun
{
    public class Weather
    {
        public int Id { get; set; }
        public string ArchiveName { get; set; }
        public DateTime? Date { get; set; }
        public double? Temp { get; set; }
        public double? Wet { get; set; }
        public double? DewPoint { get; set; }
        public double? Pressure { get; set; }
        public string WindDirect { get; set; }
        public double? WindSpeed { get; set; }
        public double? CloudCover { get; set; }
        public double? LowLimitCloud { get; set; }
        public string HorizontalVisibility { get; set; }
        public string WeatherEffect { get; set; }
    }

}