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

                var radios = this.GetRadios(document, radioRows, channels, position)
                    .Where(r => r != null)
                    .Where(r => r.ChannelToNumber != null && r.ChannelToNumber.Count > 0)
                    .ToList();

                List<ChannelModel> usedChannels = new List<ChannelModel>();
                foreach (var radio in radios)
                {
                    foreach (var c in radio.ChannelToNumber.Keys)
                    {
                        if (!usedChannels.Contains(c))
                        {
                            usedChannels.Add(c);
                        }
                    }
                }

                if (radios.Any(r => !String.IsNullOrEmpty(r.Name)))
                {
                    result.IdentifyingColumNumber = 3;
                }
                else if (radios.Any(r => !String.IsNullOrEmpty(r.Function)))
                {
                    result.IdentifyingColumNumber = 2;
                }
                else
                {
                    result.IdentifyingColumNumber = 1;
                }

                result.CourseName = GetStringValue(secondCell, document);
                result.Channels = usedChannels;
                result.Radios = radios;
            }
            return result;
        }

        private IEnumerable<RadioModel> GetRadios(SpreadsheetDocument document, IEnumerable<Row> radioRows, List<ChannelModel> channels, string channelsPosition)
        {
            return radioRows.Select(row => this.GetRadio(document, row, channels, channelsPosition));
        }

        private RadioModel GetRadio(SpreadsheetDocument document, Row radioRow, List<ChannelModel> channels,
            string channelsPosition)
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

            var inBetweenCells = radioRow.Descendants<Cell>()
                .SkipWhile(c => c != identifierCell).TakeWhile(c => !c.CellReference.Value.StartsWith(channelsPosition, StringComparison.InvariantCultureIgnoreCase))
                .ToList();

            result.Identifier = this.GetStringValue(identifierCell, document);
            Cell nameCell = null;
            Cell functionCell = null;
            Cell indicationCell = null;
            if (inBetweenCells.Count >= 4)
            {
                // name, function, indication
                nameCell = inBetweenCells[1];
                functionCell = inBetweenCells[2];
                indicationCell = inBetweenCells[3];
            }
            else if (inBetweenCells.Count >= 3)
            {
                // function and indication
                functionCell = inBetweenCells[1];
                indicationCell = inBetweenCells[2];
            }
            else if (inBetweenCells.Count >= 2)
            {
                // indication
                indicationCell = inBetweenCells[1];
            }

            if (indicationCell != null)
            {
                result.Indication = this.GetStringValue(indicationCell, document);
            }
            if (nameCell != null)
            {
                result.Name = this.GetStringValue(nameCell, document);
            }
            if (functionCell != null)
            {
                result.Function = this.GetStringValue(functionCell, document);
            }

            // now for channels
            var channelCells = radioRow.Descendants<Cell>()
                .SkipWhile(c => !c.CellReference.Value.StartsWith(channelsPosition, StringComparison.InvariantCultureIgnoreCase))
                .Take(channels.Count)
                .ToList();

            var channelToCell = channels.Zip(channelCells);

            result.ChannelToNumber = new Dictionary<ChannelModel, string>();
            foreach (var p in channelToCell)
            {
                var channel = p.First;
                var cell = p.Second;
                string positionNumber = this.GetStringValue(cell, document);
                if (!String.IsNullOrEmpty(positionNumber))
                {
                    result.ChannelToNumber[channel] = positionNumber;
                }
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
