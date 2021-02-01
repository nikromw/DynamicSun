

using System;
using System.Data.Entity;

namespace DynamicSun
{
    public class WeatherContext : DbContext
    {
        public WeatherContext()
               : base("name=WeatherContext")
        { Database.CreateIfNotExists(); }
        public DbSet<Weather> Weathers { get; set; }
        public DbSet<Archive> Archives { get; set; }

    }
}