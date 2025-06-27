using Desktop.Helpers;
using Desktop.Services.Interfaces;
using Microsoft.Office.Interop.Word;
using System.IO;
using System.Runtime.InteropServices;
using Application = Microsoft.Office.Interop.Word.Application;
using Range = Microsoft.Office.Interop.Word.Range;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;
using Office = Microsoft.Office.Core;

namespace Desktop.Services.Implementations
{
    public class InteropWordService : IInteropWordService
    {
        private Application? _wordApp;

        private async Task<Application> GetWordAppAsync()
        {
            if (_wordApp == null)
            {
                _wordApp = await Task.Run(() => new Application());
            }
            return _wordApp;
        }

        public async Task<Document> OpenDocumentAsync(string filePath, bool visible)
        {
            var wordApp = await GetWordAppAsync();
            wordApp.Visible = visible;

            return await Task.Run(() => wordApp.Documents.Open(filePath));
        }

        public async Task SaveDocumentAsync(_Document document)
        {
            try
            {
                await Task.Run(() => document.Save());
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi khi lưu tài liệu: {ex.Message}");
            }
        }

        public async Task CloseDocumentAsync(_Document document)
        {
            if (document != null)
            {
                await Task.Run(() =>
                {
                    document.Close(false);
                    ReleaseComObject(document);
                });
            }
        }

        public async Task FormatDocumentAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    float CmToPt(float cm) => (float)(cm * 28.35); // Chuyển đổi cm → pt

                    // ==== 1. Định dạng ký tự qua style "Normal"
                    var normalStyle = document.Styles["Normal"];
                    var font = normalStyle.Font;
                    font.Name = Constants.FONT_NAME;
                    font.Size = Constants.FONT_SIZE;

                    // ==== 2. Định dạng đoạn văn qua style "Normal"
                    var paraFormat = normalStyle.ParagraphFormat;
                    paraFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphJustify;
                    paraFormat.OutlineLevel = WdOutlineLevel.wdOutlineLevelBodyText;
                    paraFormat.LeftIndent = 0f;
                    paraFormat.CharacterUnitLeftIndent = 0f;
                    paraFormat.RightIndent = 0f;
                    paraFormat.SpaceBefore = 0f;
                    paraFormat.SpaceAfter = 0f;
                    paraFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;
                    paraFormat.FirstLineIndent = 0f;
                    paraFormat.HangingPunctuation = 0;
                    paraFormat.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceMultiple; // Kiểu khoảng cách dòng: Multiple
                    paraFormat.LineSpacing = (float)(1.2 * 12); // 1.2 x 12pt = 14.4pt

                    // ==== 3. Định dạng trang
                    var setup = document.PageSetup;
                    setup.PaperSize = Word.WdPaperSize.wdPaperA4;
                    setup.Orientation = Word.WdOrientation.wdOrientPortrait;
                    setup.TopMargin = CmToPt(1.27f);
                    setup.BottomMargin = CmToPt(1.27f);
                    setup.LeftMargin = CmToPt(1.27f);
                    setup.RightMargin = CmToPt(1.27f);

                    // Điều chỉnh khoảng cách Header và Footer từ lề
                    setup.HeaderDistance = CmToPt(1.27f);
                    setup.FooterDistance = CmToPt(1.27f);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
            });
        }

        public async Task FindAndReplaceAsync(_Document document, string findText, string replaceWithText, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                Word.Find findObject = null!;
                try
                {
                    findObject = document.Content.Find;
                    findObject.ClearFormatting();
                    findObject.Text = findText;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = replaceWithText;

                    findObject.MatchCase = matchCase;
                    findObject.MatchWholeWord = matchWholeWord;
                    findObject.MatchWildcards = matchWildcards;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = System.Type.Missing;

                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
                finally
                {
                    ReleaseComObject(findObject);
                }
            });

        }

        public async Task FindAndReplaceAsync(_Document document, Dictionary<string, string> replacements, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                Word.Find findObject = null!;
                try
                {
                    findObject = document.Content.Find;

                    foreach (KeyValuePair<string, string> kvp in replacements)
                    {
                        findObject.ClearFormatting();
                        findObject.Text = kvp.Key;
                        findObject.Replacement.ClearFormatting();
                        findObject.Replacement.Text = kvp.Value;

                        findObject.MatchCase = matchCase;
                        findObject.MatchWholeWord = matchWholeWord;
                        findObject.MatchWildcards = matchWildcards;

                        object replaceAll = Word.WdReplace.wdReplaceAll;
                        object missing = Type.Missing;

                        findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref missing, ref missing, ref missing, ref missing, ref missing,
                            ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
                finally
                {
                    ReleaseComObject(findObject);
                }
            });

        }

        public async Task FindAndReplaceFirstAsync(Paragraph paragraph, string findText, string replaceWithText, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                Word.Find findObject = null!;
                try
                {
                    Word.Range range = paragraph.Range;
                    findObject = range.Find;
                    findObject.ClearFormatting();
                    findObject.Text = findText;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = replaceWithText;

                    findObject.MatchCase = matchCase;
                    findObject.MatchWholeWord = matchWholeWord;
                    findObject.MatchWildcards = matchWildcards;

                    object replaceOne = Word.WdReplace.wdReplaceOne;
                    object missing = System.Type.Missing;

                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceOne, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
                finally
                {
                    ReleaseComObject(findObject);
                }
            });
        }

        public async Task FindAndReplaceInSectionAsync(_Document document, int sectionIndex, string findText, string replaceWithText, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                Word.Find findObject = null!;
                try
                {
                    if (sectionIndex < 1 || sectionIndex > document.Sections.Count)
                    {
                        MessageHelper.Error("Chỉ mục section không hợp lệ.");
                        return;
                    }

                    Word.Section section = document.Sections[sectionIndex];
                    findObject = section.Range.Find;

                    findObject.ClearFormatting();
                    findObject.Text = findText;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = replaceWithText;

                    findObject.MatchCase = matchCase;
                    findObject.MatchWholeWord = matchWholeWord;
                    findObject.MatchWildcards = matchWildcards;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;

                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
                finally
                {
                    ReleaseComObject(findObject);
                }
            });

        }

        public async Task FindAndReplaceRedToUnderlinedAsync(_Document document)
        {
            await Task.Run(() =>
            {
                Word.Find findObject = null!;
                try
                {
                    Word.Range range = document.Content;
                    findObject = range.Find;

                    findObject.ClearFormatting();
                    findObject.Font.Color = Word.WdColor.wdColorRed;
                    findObject.Text = "";

                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Font.Color = Word.WdColor.wdColorBlack;
                    findObject.Replacement.Font.Underline = Word.WdUnderline.wdUnderlineSingle;
                    findObject.Replacement.Text = "^&";

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = System.Type.Missing;
                    object format = true;

                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref format, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
                finally
                {
                    ReleaseComObject(findObject);
                }
            });
        }

        public async Task FindAndReplaceInRangeAsync(_Document document, int start, int end, string findText, string replaceWithText, bool matchCase = false, bool matchWholeWord = false, bool matchWildcards = false)
        {
            await Task.Run(() =>
            {
                Word.Find findObject = null!;
                Word.Range rangeObject = null!;
                try
                {
                    if (start < 0 || end > document.Content.End || start > end)
                    {
                        MessageHelper.Error("Giá trị start hoặc end không hợp lệ.");
                        return;
                    }

                    rangeObject = document.Range(start, end);
                    findObject = rangeObject.Find;

                    findObject.ClearFormatting();
                    findObject.Text = findText;
                    findObject.Replacement.ClearFormatting();
                    findObject.Replacement.Text = replaceWithText;

                    findObject.MatchCase = matchCase;
                    findObject.MatchWholeWord = matchWholeWord;
                    findObject.MatchWildcards = matchWildcards;

                    object replaceAll = Word.WdReplace.wdReplaceAll;
                    object missing = Type.Missing;

                    findObject.Execute(ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref missing, ref missing, ref missing, ref missing, ref missing,
                        ref replaceAll, ref missing, ref missing, ref missing, ref missing);
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
                finally
                {
                    ReleaseComObject(findObject);
                    ReleaseComObject(rangeObject);
                }
            });
        }

        public async Task ConvertListFormatToTextAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    Word.Range rng = document.Content;
                    if (rng != null && rng.ListFormat.ListType != Word.WdListType.wdListNoNumbering)
                    {
                        rng.ListFormat.ConvertNumbersToText();
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
            });
        }

        public async Task DeleteAllHeadersAndFootersAsync(_Document document)
        {
            await Task.Run(() =>
            {
                try
                {
                    foreach (Word.Section section in document.Sections)
                    {
                        foreach (Word.HeaderFooter header in section.Headers)
                        {
                            if (header.Exists)
                                header.Range.Delete();
                        }
                    }

                    foreach (Word.Section section in document.Sections)
                    {
                        foreach (Word.HeaderFooter footer in section.Footers)
                        {
                            if (footer.Exists)
                                footer.Range.Delete();
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
            });
        }

        public async Task SetAnswersToABCDAsync(_Document document)
        {
            try
            {
                int questionIndex = 0, answerIndex = 0;

                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    try
                    {
                        string str = paragraph.Range.Text.Trim();

                        if (Constants.QuestionHeaderRegex.IsMatch(str))
                        {
                            questionIndex++;
                            answerIndex = 0;
                        }

                        if (str.Contains(Constants.ANSWER_TEMPLATE) || Constants.MultipleChoiceAnswerRegex.IsMatch(str))
                        {
                            string label = GenerateLabel(answerIndex);
                            await FindAndReplaceFirstAsync(paragraph, Constants.ANSWER_TEMPLATE, $"{label} ");
                            answerIndex++;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        public async Task SetQuestionsToNumberAsync(_Document document)
        {
            try
            {
                int i = 0;
                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    try
                    {
                        string str = paragraph.Range.Text.Trim();
                        if (await IsQuestionAsync(str))
                        {
                            i++;
                            await FindAndReplaceFirstAsync(paragraph, Constants.QUESTION_TEMPLATE, $"Câu {i}: ");
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        public async Task FormatQuestionAndAnswerAsync(_Document document)
        {
            try
            {
                int soCauHoi = 0;
                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    try
                    {
                        string str = paragraph.Range.Text;
                        if (await IsQuestionAsync(str))
                        {
                            soCauHoi++;
                            if (str.StartsWith(string.Format("Câu {0}:", soCauHoi)))
                            {
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = paragraph.Range.Characters[3].Font.Bold =
                                paragraph.Range.Characters[4].Font.Bold = paragraph.Range.Characters[5].Font.Bold = paragraph.Range.Characters[6].Font.Bold = 1;
                                if (soCauHoi >= 10)
                                {
                                    paragraph.Range.Characters[7].Font.Bold = 1;
                                }
                                if (soCauHoi >= 100)
                                {
                                    paragraph.Range.Characters[8].Font.Bold = 1;
                                }
                                if (soCauHoi >= 1000)
                                {
                                    paragraph.Range.Characters[9].Font.Bold = 1;
                                }
                                if (soCauHoi >= 10000)
                                {
                                    paragraph.Range.Characters[10].Font.Bold = 1;
                                }
                            }
                            else if (str.StartsWith("<#>") || str.StartsWith("<G>") || str.StartsWith("<g>"))
                            {
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = paragraph.Range.Characters[3].Font.Bold = 1;
                            }
                            else if (str.StartsWith("<NB>") || str.StartsWith("<TH>") || str.StartsWith("<VD>"))
                            {
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold =
                                    paragraph.Range.Characters[3].Font.Bold = paragraph.Range.Characters[4].Font.Bold = 1;
                            }
                            else if (str.StartsWith("<VDC>"))
                            {
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold =
                                    paragraph.Range.Characters[3].Font.Bold = paragraph.Range.Characters[4].Font.Bold = paragraph.Range.Characters[5].Font.Bold = 1;
                            }
                            else if (str.StartsWith("#"))
                            {
                                paragraph.Range.Characters[1].Font.Bold = 1;
                            }
                            else if (str.StartsWith("[<br>]"))
                            {
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = paragraph.Range.Characters[3].Font.Bold =
                                paragraph.Range.Characters[4].Font.Bold = paragraph.Range.Characters[5].Font.Bold = paragraph.Range.Characters[6].Font.Bold = 1;
                            }
                        }
                        else if (str.StartsWith("A.") || str.StartsWith("B.") || str.StartsWith("C.") || str.StartsWith("D.") ||
                                 str.StartsWith("a)") || str.StartsWith("b)") || str.StartsWith("c)") || str.StartsWith("d)"))
                        {
                            if (paragraph.Range.Characters[1].Font.Underline == WdUnderline.wdUnderlineSingle)
                            {
                                paragraph.Range.Font.Color = WdColor.wdColorBlack;
                                paragraph.Range.Characters[1].Font.Underline = paragraph.Range.Characters[2].Font.Underline = WdUnderline.wdUnderlineSingle;
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = 1;
                                paragraph.Range.Characters[3].Font.Underline = WdUnderline.wdUnderlineNone;
                                paragraph.Range.Characters[3].Font.Bold = 0;
                            }
                            else
                            {
                                paragraph.Range.Font.Color = WdColor.wdColorBlack;
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = 1;
                                paragraph.Range.Characters[3].Font.Bold = 0;
                            }
                        }
                        else if (str.StartsWith("<$>"))
                        {

                            paragraph.Range.Font.Color = WdColor.wdColorBlack;
                            paragraph.Range.Font.Bold = 0;

                            if (paragraph.Range.Characters[1].Font.Underline == WdUnderline.wdUnderlineSingle)
                            {
                                paragraph.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                paragraph.Range.Characters[1].Font.Underline = paragraph.Range.Characters[2].Font.Underline = paragraph.Range.Characters[3].Font.Underline = WdUnderline.wdUnderlineSingle;
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = paragraph.Range.Characters[3].Font.Bold = 1;
                            }
                            else
                            {
                                paragraph.Range.Font.Underline = WdUnderline.wdUnderlineNone;
                                paragraph.Range.Characters[1].Font.Bold = paragraph.Range.Characters[2].Font.Bold = paragraph.Range.Characters[3].Font.Bold = 1;
                            }
                        }
                        else
                        {
                            paragraph.Range.Font.Color = WdColor.wdColorBlack;
                        }
                    }
                    finally
                    {
                        ReleaseComObject(paragraph);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        public async Task ProcessImagesInDocumentAsync(Word._Document document, bool isBorderImage)
        {
            try
            {
                foreach (Word.InlineShape inlineShape in document.InlineShapes)
                {
                    try
                    {
                        if (inlineShape.Type != Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject ||
                            !inlineShape.OLEFormat.ProgID.Contains("Equation"))
                        {
                            await ApplyImageBorderAsync(inlineShape);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageHelper.Error($"Lỗi xử lý hình ảnh (InlineShape): {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObject(inlineShape);
                    }
                }

                foreach (Word.Shape shape in document.Shapes)
                {
                    try
                    {
                        if (shape.Type != Office.MsoShapeType.msoEmbeddedOLEObject ||
                            !shape.OLEFormat.ProgID.Contains("Equation"))
                        {
                            await ApplyImageBorderAsync(shape);
                        }
                    }
                    catch (Exception ex)
                    {
                        MessageHelper.Error($"Lỗi xử lý hình ảnh (Shape): {ex.Message}");
                    }
                    finally
                    {
                        ReleaseComObject(shape);
                    }
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi tổng khi xử lý hình ảnh: {ex.Message}");
            }
        }

        private Task ApplyImageBorderAsync(dynamic shape)
        {
            shape.Borders.Enable = 1;
            shape.Borders.OutsideLineStyle = Word.WdLineStyle.wdLineStyleSingle;
            shape.Borders.OutsideColor = Word.WdColor.wdColorBlack;
            shape.Borders.OutsideLineWidth = Word.WdLineWidth.wdLineWidth025pt;

            return Task.CompletedTask;
        }

        public async Task<string> ConvertDocxToXpsAsync(string docxPath)
        {
            string xpsPath = Path.ChangeExtension(Path.GetTempFileName(), ".xps");
            _Document? doc = null;

            try
            {
                doc = await OpenDocumentAsync(docxPath, false); // Sử dụng hàm async
                await Task.Run(() => doc.ExportAsFixedFormat(xpsPath, WdExportFormat.wdExportFormatXPS));
            }
            catch (Exception ex)
            {
                MessageHelper.Error($"Lỗi khi chuyển đổi Docx sang XPS: {ex.Message}");
            }
            finally
            {
                if (doc != null)
                {
                    await CloseDocumentAsync(doc);
                }
            }

            return xpsPath;
        }

        public async Task UpdateFieldsAsync(string filePath)
        {
            _Document? doc = null;
            try
            {
                doc = await OpenDocumentAsync(filePath, false);
                doc.Fields.Update();
                await SaveDocumentAsync(doc);
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
            finally
            {
                if (doc != null)
                {
                    await CloseDocumentAsync(doc);
                }
            }
        }

        public async Task ClearTabStopsAsync(Paragraph paragraph)
        {
            await Task.Run(() =>
            {
                Word.TabStops tabStops = null!;

                try
                {
                    tabStops = paragraph.Format.TabStops;
                    for (int i = tabStops.Count; i >= 1; i--)
                    {
                        try
                        {
                            Word.TabStop tab = tabStops[i];
                            tabStops[tab.Position].Clear();
                            ReleaseComObject(tab);
                        }
                        catch(Exception ex)
                        {
                            MessageHelper.Error($"Lỗi ClearTabStopsAsync: {ex.Message}");
                        }
                    }
                }
                finally
                {
                    ReleaseComObject(tabStops);
                }
            });
        }

        private async Task<bool> IsQuestionAsync(string s) => await Task.Run(() => Constants.QuestionPrefixes.Any(s.StartsWith) || Constants.QuestionHeaderRegex.IsMatch(s));

        private async Task<bool> IsAnswerAsync(string s) => await Task.Run(() => Constants.AnswerPrefixes.Any(s.StartsWith));

        private string GenerateLabel(int index)
        {
            const int baseChar = 'A';
            string label = "";
            index++;

            while (index > 0)
            {
                label = (char)(baseChar + (--index % 26)) + label;
                index /= 26;
            }

            return label + ".";
        }

        private void ReleaseComObject(object? comObject)
        {
            try
            {
                if (comObject != null)
                {
                    Marshal.ReleaseComObject(comObject);
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        public async Task CloseWordAppAsync()
        {
            if (_wordApp == null) return;

            await Task.Run(() =>
            {
                try
                {
                    var docs = _wordApp.Documents;
                    int count = docs.Count;

                    // Đóng từng tài liệu từ cuối về đầu để tránh lỗi chỉ mục
                    for (int i = count; i >= 1; i--)
                    {
                        Document doc = docs[i];
                        try
                        {
                            doc.Close(false);
                        }
                        catch (Exception ex)
                        {
                            MessageHelper.Error($"Lỗi khi đóng document: {ex.Message}");
                        }
                        finally
                        {
                            ReleaseComObject(doc);
                        }
                    }

                    _wordApp.Quit();
                }
                catch (Exception ex)
                {
                    MessageHelper.Error($"Lỗi khi đóng ứng dụng Word: {ex.Message}");
                }
                finally
                {
                    ReleaseComObject(_wordApp);
                    _wordApp = null;
                }
            });
        }

        public async ValueTask DisposeAsync()
        {
            await CloseWordAppAsync();

            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
    }
}
