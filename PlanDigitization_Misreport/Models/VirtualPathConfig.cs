using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;

namespace LessonLearntPortalWeb.Models
{
    public class VirtualPathConfig
    {
        public List<PathConfig> Paths { get; set; }
    }

    public class PathConfig
    {
        public string RealPath { get; set; }
        public string RequestPath { get; set; }
        public string Alias { get; set; }
    }
}
