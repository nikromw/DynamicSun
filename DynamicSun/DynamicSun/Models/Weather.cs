﻿using System;

namespace DynamicSun
{
    public class Weather
    {
        public int Id { get; set; }
        public string ArchiveName { get; set; }
        public DateTime Date { get; set; }
        public double Temp { get; set; }
        public int Wet { get; set; }
        public double DewPoint { get; set; }
        public int Pressure { get; set; }
        public string WindDirect { get; set; }
        public double WindSpeed { get; set; }
        public double CloudCover { get; set; }
        public double LowLimitCloud { get; set; }
        public double HorizontalVisibility { get; set; }
        public string WeatherEffect { get; set; }
    }

}