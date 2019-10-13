using System;
using CourseRadioPlan.Models;
using Microsoft.AspNetCore.Http;

namespace CourseRadioPlan.Services
{
    public interface IRadioPlanSheetService
    {
        RadioPlanModel GenerateFromFile(IFormFile formFile);
    }
}
