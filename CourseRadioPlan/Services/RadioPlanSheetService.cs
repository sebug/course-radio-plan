using System;
using CourseRadioPlan.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using System.Linq;

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
            using (var document = SpreadsheetDocument.Open(formFile.OpenReadStream(), false))
            {
                var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>();
                if (!sheets.Any())
                {
                    throw new Exception("No worksheet found");
                }
                var firstSheet = sheets.First();

                WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(firstSheet.Id);
                Worksheet worksheet = worksheetPart.Worksheet;

                var fourthRow = worksheet.Descendants<Row>().Skip(3).First();

                var secondCell = fourthRow.Descendants<Cell>().Skip(1).First();

                result.CourseName = GetStringValue(secondCell, document);
            }
            return result;
        }

        private string GetStringValue(Cell c, SpreadsheetDocument document)
        {
            // If the content of the first cell is stored as a shared string, get the text of the first cell
            // from the SharedStringTablePart and return it. Otherwise, return the string value of the cell.
            if (c.DataType != null && c.DataType.Value ==
                CellValues.SharedString)
            {
                SharedStringTablePart shareStringPart = document.WorkbookPart.
            GetPartsOfType<SharedStringTablePart>().First();
                SharedStringItem[] items = shareStringPart.
            SharedStringTable.Elements<SharedStringItem>().ToArray();
                return items[int.Parse(c.CellValue.Text)].InnerText;
            }
            else
            {
                return c.CellValue.Text;
            }
        }
    }
}
