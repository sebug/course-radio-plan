using System;
using Microsoft.AspNetCore.Http;

namespace CourseRadioPlan.Services
{
    public class RadioPlanSheetService : IRadioPlanSheetService
    {
        public RadioPlanSheetService()
        {
        }

        public string GenerateFromFile(IFormFile formFile)
        {
            return "Course name already";
        }
    }
}
