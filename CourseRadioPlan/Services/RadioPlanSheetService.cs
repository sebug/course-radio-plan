using System;
using CourseRadioPlan.Models;
using Microsoft.AspNetCore.Http;

namespace CourseRadioPlan.Services
{
    public class RadioPlanSheetService : IRadioPlanSheetService
    {
        public RadioPlanSheetService()
        {
        }

        public RadioPlanModel GenerateFromFile(IFormFile formFile)
        {
            var result = new RadioPlanModel
            {
                CourseName = "The course name"
            };
            return result;
        }
    }
}
