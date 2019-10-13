using System;
using Microsoft.AspNetCore.Http;

namespace CourseRadioPlan.Services
{
    public interface IRadioPlanSheetService
    {
        string GenerateFromFile(IFormFile formFile);
    }
}
