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
            return "<html>" +
                "<head>" +
                "<title>The Sheet</title>" +
                "</head>" +
                "<body>" +
                "<h1>Ohai, world</h1>" +
                "</body>" +
                "</html>";
        }
    }
}
