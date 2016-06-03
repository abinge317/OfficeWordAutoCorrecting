using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml.Spreadsheet;

namespace OfficeWordAutoCorrecting
{
    public class WordAutoCorrectingTool
    {
        private static SpreadsheetDocument spreadsheetDocument = null;
        public const int PARAGRAPH_BORDER_TESTPOINT_ROW = 1;
        public const int PARAGRAPH_TEXT1_ROW = 1;
        public const int PARAGRAPH_TEXT1_COL = 8;

        /// <summary>
        /// 调用该方法即可进行自动批改操作
        /// </summary>
        /// <param name="testPointFilePath">定义考察点的excel文件所在路径，该文件需要定义此次操作题所涉及到的考察点以及对应的分值，必须满足规定格式</param>
        /// <param name="standardAnswerFilePath">针对于此次考试的标准答案文件所在路径，批改操作完全按照该文件进行，必须保证该标准答案百分之百正确</param>
        /// <param name="studentFilePath">需要进行批改的学生的Word文件所在路径</param>
        /// <returns>所得分数</returns>
        public static int DoCorrect(string testPointFilePath, string standardAnswerFilePath, string studentFilePath)
        {
            InitialTestPoint(testPointFilePath);
            InitialStandardAnswer(standardAnswerFilePath);
            int score = StartCorrect(studentFilePath);
            return score;
        }

        /// <summary>
        /// 初始化考点信息，获取考点和对应分值之间的对应关系
        /// </summary>
        /// <param name="testPointFilePath">定义考察点的excel文件所在路径</param>
        /// <returns></returns>
        private static void InitialTestPoint(string testPointFilePath)
        {
            spreadsheetDocument = SpreadsheetDocument.Open(testPointFilePath, false);
            IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            Row[] rows = Helper.getRowsBySheetName(spreadsheetDocument, sheets, ConstantValue.TEMPLATE_FILE_SHEET1_NAME);
            int score = int.Parse(Helper.GetCellValue(spreadsheetDocument, rows[PARAGRAPH_BORDER_TESTPOINT_ROW].Descendants<Cell>().ElementAt(1)));
            if (Helper.CellValueEquals(spreadsheetDocument, rows, PARAGRAPH_BORDER_TESTPOINT_ROW, 0, ConstantValue.BORDER_TYPE_STRING))
            {
                if (score != 0) TestPoint.testPoint2Score[TestPoint.BORDER_TYPE] = score;
            }
            else throw new FormatException();

            score = int.Parse(Helper.GetCellValue(spreadsheetDocument, rows[PARAGRAPH_BORDER_TESTPOINT_ROW].Descendants<Cell>().ElementAt(3)));
            if (Helper.CellValueEquals(spreadsheetDocument, rows, PARAGRAPH_BORDER_TESTPOINT_ROW, 2, ConstantValue.BORDER_COLOR_STRING))
            {
                if (score != 0) TestPoint.testPoint2Score[TestPoint.BORDER_COLOR] = score;
            }
            else throw new FormatException();

            string cellValue = Helper.GetCellValue(spreadsheetDocument, rows[PARAGRAPH_BORDER_TESTPOINT_ROW].Descendants<Cell>().ElementAt(5));
            try
            {
                score = int.Parse(cellValue);
            }
            catch (FormatException)
            {
                throw new FormatException("Excel文件分值格式错误");
            }
            if (Helper.CellValueEquals(spreadsheetDocument, rows, PARAGRAPH_BORDER_TESTPOINT_ROW, 4, ConstantValue.BORDER_SIZE_STRING))
            {
                if (score != 0) TestPoint.testPoint2Score[TestPoint.BORDER_SIZE] = score;
            }
            else throw new FormatException();

            score = int.Parse(Helper.GetCellValue(spreadsheetDocument, rows[PARAGRAPH_BORDER_TESTPOINT_ROW].Descendants<Cell>().ElementAt(7)));
            if (Helper.CellValueEquals(spreadsheetDocument, rows, PARAGRAPH_BORDER_TESTPOINT_ROW, 6, ConstantValue.BORDER_SHADOW_STRING))
            {
                if (score != 0) TestPoint.testPoint2Score[TestPoint.BORDER_SHADOW] = score;
            }
            else throw new FormatException();

            //TODO 其它Sheet的初始化工作


            //spreadsheetDocument.Close();

        }

        /// <summary>
        /// 获取标准答案涉及到的各个考察点的属性值
        /// </summary>
        /// <param name="standardAnswerFilePath">针对于此次考试的标准答案文件所在路径</param>
        /// <returns></returns>
        private static void InitialStandardAnswer(string standardAnswerFilePath)
        {
            WordprocessingDocument wordprocessingDocument = WordprocessingDocument.Open(standardAnswerFilePath, false);
            IEnumerable<Paragraph> paraList = wordprocessingDocument.MainDocumentPart.Document.Descendants<Paragraph>();
            IEnumerator<Paragraph> paragraphsEnumerator = paraList.GetEnumerator();

            IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            Row[] rows = Helper.getRowsBySheetName(spreadsheetDocument, sheets, ConstantValue.TEMPLATE_FILE_SHEET1_NAME);
            string text = Helper.GetCellValue(spreadsheetDocument, rows[PARAGRAPH_TEXT1_ROW].Descendants<Cell>().ElementAt(PARAGRAPH_TEXT1_COL)).Trim();
            List<Paragraph> paragraphs = Helper.getParagraphsWithBorderAndMatchText(paragraphsEnumerator, text);
            if (paragraphs.Count < 1)
            {
                throw new FormatException("Excel文件格式有误，可能是Excel文件中指定的段落文本没有与标准答案中的段落未文本完全匹配");
            }
            Paragraph paragraph = paragraphs[0];
         
            ParagraphEntity paragraphEntity = new ParagraphEntity();
            ParagraphProperties pPro = paragraph.ParagraphProperties;
            if (pPro != null)
            {
                ParagraphBorders borders = pPro.ParagraphBorders;
                if (borders != null)
                {
                    paragraphEntity.WithBorder = true;
                    if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_SIZE))
                        paragraphEntity.BorderSize = borders.TopBorder.Size;
                    if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_TYPE))
                        paragraphEntity.BorderType = borders.TopBorder.Val.ToString();
                    if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_COLOR))
                        paragraphEntity.TopBorderColor = borders.TopBorder.Color;
                    if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_SHADOW))
                        paragraphEntity.WithShadow = borders.TopBorder.Shadow.Value;
                }
            }
            StandardAnswer.paragraphEntities[text] = paragraphEntity;

            //TODO 获取标准答案中其它考点的属性值


            wordprocessingDocument.Close();
        }

        /// <summary>
        /// 根据获取到的信息进行改卷
        /// </summary>
        /// <param name="studentFilePath">学生文件所在路径</param>
        /// <returns>分数</returns>
        private static int StartCorrect(string studentFilePath)
        {
            int score = 0;
            WordprocessingDocument studentWordprocessingDocument = WordprocessingDocument.Open(studentFilePath, false);
            IEnumerable<Paragraph> stuParaList = studentWordprocessingDocument.MainDocumentPart.Document.Descendants<Paragraph>();
            IEnumerator<Paragraph> stuParagraphsEnumerator = stuParaList.GetEnumerator();
            IEnumerable<Sheet> sheets = spreadsheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
            Row[] rows = Helper.getRowsBySheetName(spreadsheetDocument, sheets, ConstantValue.TEMPLATE_FILE_SHEET1_NAME);
            string text = Helper.GetCellValue(spreadsheetDocument, rows[PARAGRAPH_TEXT1_ROW].Descendants<Cell>().ElementAt(PARAGRAPH_TEXT1_COL)).Trim();
            List<Paragraph> stuParagraphs = Helper.getParagraphsWithBorderAndMatchText(stuParagraphsEnumerator, text);
            if (stuParagraphs.Count == 1)
            {
                ParagraphEntity paragraphEntity = StandardAnswer.paragraphEntities[text];

                ParagraphProperties pPro = stuParagraphs[0].ParagraphProperties;
                if (pPro != null)
                {
                    ParagraphBorders borders = pPro.ParagraphBorders;
                    if (borders != null)
                    {
                        if (borders.TopBorder != null)
                        {
                            if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_SIZE) && paragraphEntity.BorderSize == borders.TopBorder.Size)
                            {
                                //logTextbox.AppendText("边框粗细正确, +" + testPoint2Score[ParagraphTestPoint.BORDER_SIZE] + "分\r\n");
                                score += TestPoint.testPoint2Score[TestPoint.BORDER_SIZE];

                            }
                            if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_TYPE) && paragraphEntity.BorderType == borders.TopBorder.Val.ToString())
                            {
                                //logTextbox.AppendText("边框类型正确, +" + testPoint2Score[ParagraphTestPoint.BORDER_TYPE] + "分\r\n");
                                score += TestPoint.testPoint2Score[TestPoint.BORDER_TYPE];
                            }
                            if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_COLOR) && paragraphEntity.TopBorderColor == borders.TopBorder.Color)
                            {
                                //logTextbox.AppendText("边框颜色正确, +" + testPoint2Score[ParagraphTestPoint.BORDER_COLOR] + "分\r\n");
                                score += TestPoint.testPoint2Score[TestPoint.BORDER_COLOR];
                            }
                            if (TestPoint.testPoint2Score.ContainsKey(TestPoint.BORDER_SHADOW) && (borders.TopBorder.Shadow != null && paragraphEntity.WithShadow == borders.TopBorder.Shadow.Value))
                            {
                                //logTextbox.AppendText("边框阴影正确, +" + testPoint2Score[ParagraphTestPoint.BORDER_SHADOW] + "分\r\n");
                                score += TestPoint.testPoint2Score[TestPoint.BORDER_SHADOW];
                            }
                        }
                    }
                }
            }
            studentWordprocessingDocument.Close();
            spreadsheetDocument.Close();
            return score;

            //TODO 批改其它地方
        }

    }
}
