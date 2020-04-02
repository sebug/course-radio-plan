using System;
using System.Collections.Generic;

namespace CourseRadioPlan.Models
{
    public class RadioPlanModel
    {
        public string CourseName { get; set; }
        public List<ChannelModel> Channels { get; set; }
        public List<RadioModel> Radios { get; set; }
        public int IdentifyingColumNumber { get; set; }
        public bool UseSVG { get; set; }
    }
}
