using System;
using System.Collections.Generic;

namespace CourseRadioPlan.Models
{
    public class RadioModel
    {
        public string Identifier { get; set; }
        public string Name { get; set; }
        public string Function { get; set; }
        public string Indication { get; set; }
        public Dictionary<ChannelModel, string> ChannelToNumber { get; set; }
    }
}
