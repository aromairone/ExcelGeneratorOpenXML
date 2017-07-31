using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.IO;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

namespace ExcelGeneratorOpenXML
{
    public class ExcelService
    {
        // The stream that the spreadsheet gets returned on
        public MemoryStream SpreadsheetStream { get; set; }

        private Worksheet CurrentWorkSheet { get { return _spreadsheet.WorkbookPart.WorksheetParts.First().Worksheet; } }
        private SpreadsheetDocument _spreadsheet;
        private Columns _columns;

        public bool CreateSpreadSheet()
        {
            //Initiate the stream
            SpreadsheetStream = new MemoryStream();
            try
            {
                //Create the spreadsheet in memory
                _spreadsheet = SpreadsheetDocument.Create(SpreadsheetStream, SpreadsheetDocumentType.Workbook);

                WorkbookPart workbookPart = _spreadsheet.AddWorkbookPart(); //Add WorkbookPart

                Workbook workbook = new Workbook(); //Initiate workbook

                Worksheet worksheet = new Worksheet(); //Initiate WorkSheet

                SheetData sheetData = new SheetData(); //Initiate SheetData

                WorkbookStylesPart stylesPart = _spreadsheet.WorkbookPart.AddNewPart<WorkbookStylesPart>();
                                
                stylesPart.Stylesheet = GenerateStyleSheet(); //Add StyleSheet
                stylesPart.Stylesheet.Save();
                
                worksheet.Append(sheetData);

                _spreadsheet.WorkbookPart.Workbook = workbook;
                _spreadsheet.WorkbookPart.Workbook.Save();
            }
            catch
            {
                return false;
            }
            return true;
        }
        private Stylesheet GenerateStyleSheet()
        {
            return new Stylesheet(
                new Fonts(
                    new Font(                                                               // Index 0 - The default font.
                        new FontSize() { Val = 8 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" }),
                    new Font(                                                               // Index 1 - Sheet TITLE The bold font.
                        new Bold(),
                        new FontSize() { Val = 12 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Arial" }),
                    new Font(                                                               // Index 2 - The HEADER "Bold/Arial/White color/8".
                        new Bold(),
                        new FontSize() { Val = 8 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "FFFFFF" } },
                        new FontName() { Val = "Arial" }),
                    new Font(                                                               // Index 3 - The Times Roman font. with 16 size
                        new FontSize() { Val = 16 },
                        new Color() { Rgb = new HexBinaryValue() { Value = "000000" } },
                        new FontName() { Val = "Times New Roman" })

                 ),//Fonts
                new Fills(
                    new Fill(                                                           // Index 0 - The default fill.
                        new PatternFill() { PatternType = PatternValues.None }),
                    new Fill(                                                           // Index 1 - The default fill of gray 125 (required)
                        new PatternFill() { PatternType = PatternValues.Gray125 }),
                    new Fill(                                                           // Index 2 - The yellow fill.
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "FFFFFF00" } }
                            )
                        { PatternType = PatternValues.Solid }),
                    new Fill(                                                           // Index 3 - The Blue fill for HEADER.
                        new PatternFill(
                                new ForegroundColor() { Rgb = new HexBinaryValue() { Value = "0000FF" } }
                            )
                        { PatternType = PatternValues.Solid })
                    ), //Fills
                new Borders(
                    new Border(                                                         // Index 0 - The default border.
                        new LeftBorder(),
                        new RightBorder(),
                        new TopBorder(),
                        new BottomBorder(),
                        new DiagonalBorder()),
                    new Border(                                                         // Index 1 - Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new RightBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new TopBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new BottomBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new DiagonalBorder()),
                    new Border(                                                         // Index 2 - Applies a Left, Right, Top, Bottom border to a cell
                        new LeftBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new RightBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new TopBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new BottomBorder(
                                new Color() { Auto = true }
                            )
                        { Style = BorderStyleValues.Medium },
                        new DiagonalBorder())
                    ), //Borders
                    new CellFormats(
                        new CellFormat() { FontId = 0, FillId = 0, BorderId = 2, ApplyBorder = true },      // Index 0 - BODY - The default cell style.  If a cell does not have a style index applied it will use this style combination instead
                        new CellFormat(
                                new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                            )
                        { FontId = 1, FillId = 0, BorderId = 0, ApplyFont = true },                   // Index 1 - Sheet TITLE
                        new CellFormat() { FontId = 2, FillId = 0, BorderId = 1, ApplyFont = true },       // Index 2 - Italic
                        new CellFormat() { FontId = 3, FillId = 2, BorderId = 1, ApplyFont = true },       // Index 3 - Times Roman 16/
                        new CellFormat() { FontId = 0, FillId = 2, BorderId = 1, ApplyFill = true },       // Index 4 - Yellow Fill
                        new CellFormat(                                                                    // Index 5 - HEADER - Alignment center /Font Bold / Blue Background/ Border
                                new Alignment() { Horizontal = HorizontalAlignmentValues.Center, Vertical = VerticalAlignmentValues.Center, WrapText = true }
                            )
                        { FontId = 2, FillId = 3, BorderId = 1, ApplyAlignment = true, ApplyBorder = true },
                        new CellFormat() { FontId = 0, FillId = 0, BorderId = 0, NumberFormatId = UInt32Value.FromUInt32(3453), ApplyNumberFormat = true }      // Index 6 - Number
                    )
            ); //Return StyleSheet
        }
    }
}
