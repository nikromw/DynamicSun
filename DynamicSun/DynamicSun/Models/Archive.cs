using System;
using System.Collections.Generic;

namespace DynamicSun
{
    public class Archive
    {
        public int Id { get; set; }
        public string ArchiveName { get; set; }
        public List<Weather> weathers { get; set; }
    }

}