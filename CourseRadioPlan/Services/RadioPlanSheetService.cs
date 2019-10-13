using System;
using CourseRadioPlan.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using System.Linq;
using System.Collections;
using System.Collections.Generic;

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

                var channels = this.GetChannels(document,
                    worksheet.Descendants<Row>().Skip(13).First(),
                    worksheet.Descendants<Row>().Skip(14).First());

                result.CourseName = GetStringValue(secondCell, document);
            }
            return result;
        }

        private IEnumerable<ChannelModel> GetChannels(SpreadsheetDocument document, Row typeRow, Row identifierRow)
        {
            var typeCells = typeRow.Descendants<Cell>().Skip(7).ToList();
            var identifierCells = identifierRow.Descendants<Cell>().Skip(8).ToList(); // skip one more since
            // one cell above was merged

            return typeCells.Zip(identifierCells, (typeCell, identifierCell) => new ChannelModel
            {
                Type = this.GetStringValue(typeCell, document),
                Number = GetStringValue(identifierCell, document)
            }).Where(cm => !String.IsNullOrEmpty(cm.Number)).ToList();
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
                if (c.CellValue != null)
                {
                    return c.CellValue.Text;
                }
                else
                {
                    return c.InnerText;
                }
            }
        }
    }
}
