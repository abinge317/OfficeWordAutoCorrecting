using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeWordAutoCorrecting
{
    class Helper
    {
        /// <summary>
        /// 获取具有边框的段落
        /// </summary>
        /// <param name="paragraphs">所有段落的迭代器</param>
        /// <returns>所有具有边框的段落列表</returns>
        [Obsolete("This function is obsolete", true)]
        public static List<Paragraph> getParagraphsWithBorder(IEnumerator<Paragraph> paragraphs)
        {
            List<Paragraph> paras = new List<Paragraph>();
            while (paragraphs.MoveNext())
            {
                ParagraphProperties pPro = paragraphs.Current.ParagraphProperties;
                if (pPro != null && pPro.ParagraphBorders != null)
                {
                    paras.Add(paragraphs.Current);
                }
            }
            return paras;
        }

        /// <summary>
        /// 获取具有边框并且与指定文本匹配的段落
        /// </summary>
        /// <param name="paragraphs">所有段落的迭代器</param>
        /// <param name="paragraphs">将要匹配的段落文本</param>
        /// <returns>所有具有边框并且与指定文本匹配的段落列表</returns>
        public static List<Paragraph> getParagraphsWithBorderAndMatchText(IEnumerator<Paragraph> paragraphs, string text)
        {
            List<Paragraph> paras = new List<Paragraph>();
            while (paragraphs.MoveNext())
            {
                ParagraphProperties pPro = paragraphs.Current.ParagraphProperties;
                Paragraph paragraph = paragraphs.Current;
                string paragraphText = paragraph.InnerText.Trim();

                if (pPro != null && pPro.ParagraphBorders != null && paragraphText == text)
                {
                    paras.Add(paragraphs.Current);
                }
            }
            return paras;
        }

        /// <summary>
        /// 根据Sheet的名字获取该Sheet所有行的数组
        /// </summary>
        /// <param name="spreadsheetDocument"></param>
        /// <param name="sheets"></param>
        /// <param name="sheetName">Sheet名称</param>
        /// <returns>该Sheet的所有行</returns>
        public static Row[] getRowsBySheetName(SpreadsheetDocument spreadsheetDocument, IEnumerable<Sheet> sheets, string sheetName)
        {
            string relationshipId = sheets.First().Id.Value = sheets.First(x => x.Name == sheetName).Id.Value;
            WorksheetPart worksheetPart = (WorksheetPart)spreadsheetDocument.WorkbookPart.GetPartById(relationshipId);
            Worksheet workSheet = worksheetPart.Worksheet;
            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
            Row[] rows = sheetData.Descendants<Row>().ToArray();
            return rows;
        }

        public static bool CellValueEquals(SpreadsheetDocument spreadsheetDocument, Row[] rows, int rowIndex, int colIndex, string value)
        {
            return GetCellValue(spreadsheetDocument, rows[rowIndex].Descendants<Cell>().ElementAt(colIndex)) == value;
        }

        public static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (value.Trim() == "")
            {
                return "0";
            }

            if (cell.DataType != null && (cell.DataType.Value == CellValues.SharedString || cell.DataType.Value == CellValues.String || cell.DataType.Value == CellValues.Number))
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else //浮点数和日期对应的cell.DataType都为NULL
            {
                // DateTime.FromOADate((double.Parse(value)); 如果确定是日期就可以直接用过该方法转换为日期对象，可是无法确定DataType==NULL的时候这个CELL 数据到底是浮点型还是日期.(日期被自动转换为浮点
                return value;
            }
        }

    }

}
