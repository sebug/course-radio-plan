using System;
using CourseRadioPlan.Models;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using Microsoft.AspNetCore.Http;
using System.Linq;
using System.Collections;
using System.Collections.Generic;
using System.Text.RegularExpressions;

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

                var positionRow = worksheet.Descendants<Row>().Skip(12).First();
                var typeRow = worksheet.Descendants<Row>().Skip(13).First();
                var numberRow = worksheet.Descendants<Row>().Skip(14).First();

                var positionAndChannelsTuple = this.GetChannels(document,
                    positionRow,
                    typeRow,
                    numberRow);

                string position = positionAndChannelsTuple.Item1;
                Regex numberRegex = new Regex("\\d+");
                position = numberRegex.Replace(position, String.Empty);
                var channels = positionAndChannelsTuple.Item2.ToList();

                var radioRows = worksheet.Descendants<Row>().Skip(15);

                var radios = this.GetRadios(document, radioRows, channels)
                    .Where(r => r != null)
                    .ToList();



                result.CourseName = GetStringValue(secondCell, document);
            }
            return result;
        }

        private IEnumerable<RadioModel> GetRadios(SpreadsheetDocument document, IEnumerable<Row> radioRows, List<ChannelModel> channels)
        {
            return radioRows.Select(row => this.GetRadio(document, row, channels));
        }

        private RadioModel GetRadio(SpreadsheetDocument document, Row radioRow, List<ChannelModel> channels)
        {
            var result = new RadioModel();

            var identifierCell = radioRow.Descendants<Cell>()
                .Where(c =>
                {
                    string v = this.GetStringValue(c, document);
                    if (v != null && v.Length == 4)
                    {
                        return true;
                    }
                    return false;
                }).FirstOrDefault();

            if (identifierCell == null)
            {
                return null;
            }

            return result;
        }

        private Tuple<string, IEnumerable<ChannelModel>> GetChannels(SpreadsheetDocument document, Row positionRow, Row typeRow, Row identifierRow)
        {
            var firstParseablePositionCell =
                positionRow.Descendants<Cell>().FirstOrDefault(c =>
                {
                    string t = this.GetStringValue(c, document);
                    int res;
                    return int.TryParse(t, out res);
                });

            if (firstParseablePositionCell == null)
            {
                throw new Exception("Did not find the first position cell");
            }
            int idx = positionRow.Descendants<Cell>().ToList().IndexOf(firstParseablePositionCell);

            var positionCells = positionRow.Descendants<Cell>().Skip(idx).ToList();
            var typeCells = typeRow.Descendants<Cell>().Skip(idx + 1).ToList();
            var identifierCells = identifierRow.Descendants<Cell>().Skip(idx + 2).ToList(); // skip one more since
            // one cell above was merged

            var positionAndTypes = positionCells.Zip(typeCells);


            return new Tuple<string, IEnumerable<ChannelModel>>(firstParseablePositionCell.CellReference.Value, positionAndTypes.Zip(identifierCells, (typeAndPosition, identifierCell) => new ChannelModel
            {
                Position = this.GetStringValue(typeAndPosition.First, document),
                Type = this.GetStringValue(typeAndPosition.Second, document),
                Number = GetStringValue(identifierCell, document)
            }).Where(cm => !String.IsNullOrEmpty(cm.Number)).ToList());
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
