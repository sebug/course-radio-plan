using System;
using Microsoft.AspNetCore.Http;

namespace CourseRadioPlan.Models
{
    public class HomeViewModel
    {
        public IFormFile ExcelFile { get; set; }

        public string CourseName { get; set; }
    }
}
