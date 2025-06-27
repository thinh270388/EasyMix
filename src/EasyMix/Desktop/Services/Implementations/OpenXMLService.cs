using Desktop.Helpers;
using Desktop.Models;
using Desktop.Models.Enums;
using Desktop.Services.Interfaces;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using System.IO;
using System.Text;
using System.Text.RegularExpressions;
using Break = DocumentFormat.OpenXml.Wordprocessing.Break;
using Level = Desktop.Models.Enums.Level;
using Paragraph = DocumentFormat.OpenXml.Wordprocessing.Paragraph;
using ParagraphProperties = DocumentFormat.OpenXml.Wordprocessing.ParagraphProperties;
using Path = System.IO.Path;
using Run = DocumentFormat.OpenXml.Wordprocessing.Run;
using RunProperties = DocumentFormat.OpenXml.Wordprocessing.RunProperties;
using SectionProperties = DocumentFormat.OpenXml.Wordprocessing.SectionProperties;
using Table = DocumentFormat.OpenXml.Wordprocessing.Table;
using TableCell = DocumentFormat.OpenXml.Wordprocessing.TableCell;
using TableRow = DocumentFormat.OpenXml.Wordprocessing.TableRow;
using TabStop = DocumentFormat.OpenXml.Wordprocessing.TabStop;
using Text = DocumentFormat.OpenXml.Wordprocessing.Text;

namespace Desktop.Services.Implementations
{
    public class OpenXMLService : IOpenXMLService
    {
        private readonly IInteropWordService _interopWordService;
        private readonly IExcelAnswerExporter _excelAnswerExporter;

        public OpenXMLService(IInteropWordService interopWordService, IExcelAnswerExporter excelAnswerExporter) 
        {
            _interopWordService = interopWordService;
            _excelAnswerExporter = excelAnswerExporter;
        }

        public Task<WordprocessingDocument> CreateDocumentAsync(string filePath)
        {
            var document = WordprocessingDocument.Create(filePath, WordprocessingDocumentType.Document);
            document.AddMainDocumentPart();
            document.MainDocumentPart!.Document = new Document(new Body());
            document.MainDocumentPart.Document.Save();
            return Task.FromResult(document);
        }

        public Task<WordprocessingDocument> OpenDocumentAsync(string filePath, bool isEditTable)
        {
            return Task.FromResult(WordprocessingDocument.Open(filePath, isEditTable));
        }

        public Task SaveDocumentAsync(WordprocessingDocument document)
        {
            document.MainDocumentPart!.Document.Save();
            return Task.CompletedTask;
        }

        public Task CloseDocumentAsync(WordprocessingDocument document)
        {
            document.Dispose();
            return Task.CompletedTask;
        }

        public Task<string> GetDocumentTextAsync(WordprocessingDocument document)
        {
            return Task.FromResult(document.MainDocumentPart!.Document.Body!.InnerText);
        }

        public Task<Body> GetDocumentBodyAsync(WordprocessingDocument document)
        {
            return Task.FromResult(document.MainDocumentPart!.Document.Body!);
        }

        public Task FormatAllParagraphsAsync(WordprocessingDocument doc)
        {
            var body = doc.MainDocumentPart!.Document.Body!;
            foreach (var para in body.Descendants<Paragraph>())
                FormatParagraph(para);
            return Task.CompletedTask;
        }

        public async Task GenerateShuffledExamsAsync(string sourceFile, string outputFolder, MixInfo mixInfo)
        {
            await Task.Run(async () =>
            {
                string mixTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Constants.TEMPLATES_FOLDER, Constants.MIX_TEMPLATE_FILE);
                string guideTemplate = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, Constants.TEMPLATES_FOLDER, Constants.GUIDE_TEMPLATE_FILE);

                if (Directory.Exists(outputFolder))
                    Directory.Delete(outputFolder, true);
                Directory.CreateDirectory(outputFolder);

                string answerFolder = Path.Combine(outputFolder, Constants.ANSWERS_FOLER);
                Directory.CreateDirectory(answerFolder);

                var allAnswers = new List<QuestionExport>();

                foreach (var version in mixInfo.Versions!)
                {
                    string mixFile = Path.Combine(outputFolder, $"{Constants.EXAM_PREFIX}{version}.docx");
                    File.Copy(sourceFile, mixFile, true);
                    string answerFile = Path.Combine(answerFolder, $"{Constants.ANSWER_PREFIX}{version}.docx");
                    File.Copy(guideTemplate, answerFile, true);

                    using (var mixDoc = WordprocessingDocument.Open(mixFile, true))
                    using (var answerDoc = WordprocessingDocument.Open(answerFile, true))
                    {
                        // Truyền các đối tượng tài liệu để tránh mở/đóng nhiều lần
                        var answers = await ShuffleQuestionsAsync(mixDoc, version, answerDoc);

                        await InsertTemplateAsync(mixTemplate, mixDoc, mixInfo, version);
                        foreach (var a in answers)
                        {
                            a.Version = version;
                            allAnswers.Add(a);
                        }
                        await AddFooterAsync(mixDoc, version);
                        await AddEndNotesAsync(mixDoc);

                        await AppendGuideAsync(answerDoc, answers, mixInfo, version);
                        await MoveEssayTableToEndAsync(answerDoc);

                        await FormatAllParagraphsAsync(answerDoc);

                        mixDoc.MainDocumentPart!.Document.Save();
                        answerDoc.MainDocumentPart!.Document.Save();
                    }
                    await _interopWordService.UpdateFieldsAsync(mixFile);
                }
                _excelAnswerExporter.ExportExcelAnswers($"{outputFolder}\\{Constants.EXCEL_ANSWER_FILE}", allAnswers);
            });
        }

        private async Task<List<QuestionExport>> ShuffleQuestionsAsync(WordprocessingDocument doc, string version, WordprocessingDocument? answerDoc = null)
        {
            return await Task.Run(async () =>
            {
                var body = doc.MainDocumentPart!.Document.Body;

                var allBlocks = SplitQuestions(body!);
                var grouped = new Dictionary<QuestionType, List<List<OpenXmlElement>>>
            {
                { QuestionType.MultipleChoice, new() },
                { QuestionType.TrueFalse, new() },
                { QuestionType.ShortAnswer, new() },
                { QuestionType.Essay, new() }
            };

                foreach (var block in allBlocks)
                {
                    var type = DetectQuestionType(block);
                    grouped[type].Add(block);
                }

                var rng = new Random();
                foreach (var key in grouped.Keys.ToList())
                {
                    if (version.Equals(Constants.ROOT_CODE)) continue;
                    grouped[key] = grouped[key].OrderBy(_ => rng.Next()).ToList();
                }

                var answers = new List<QuestionExport>();
                body!.RemoveAllChildren();
                int questionNumber = 0;
                int index = 0;

                foreach (var group in grouped.OrderBy(g => g.Key).Where(g => g.Value.Any()))
                {
                    int localQuestion = 0;
                    index++;

                    string title = CreateSectionTitle(group.Key, index, group.Value.Count);
                    if (!string.IsNullOrEmpty(title))
                    {
                        var heading = new Paragraph();
                        var parts = title.Split('.', 3);
                        if (parts.Length >= 3)
                        {
                            var boldRun = new Run(new RunProperties(new Bold()), new Text($"{parts[0]}.{parts[1]}.") { Space = SpaceProcessingModeValues.Preserve });
                            var normalRun = new Run(new Text(parts[2]) { Space = SpaceProcessingModeValues.Preserve });
                            heading.Append(boldRun, normalRun);
                        }
                        body.Append(heading);
                    }

                    foreach (var block in group.Value)
                    {
                        localQuestion++;
                        questionNumber++;
                        var type = group.Key;
                        var newBlock = ShuffleAnswers(block, type, version, doc.MainDocumentPart!, out string correct, out var answerElements);

                        // Cập nhật số thứ tự câu hỏi hiển thị
                        var firstPara = newBlock.OfType<Paragraph>().FirstOrDefault();
                        if (firstPara != null)
                        {
                            await UpdateQuestionNumberAsync(firstPara, localQuestion);
                        }
                        // Lấy điểm cho câu hỏi tự luận
                        string? point = null;
                        if (type == QuestionType.Essay)
                        {
                            // Tìm điểm trong toàn bộ block (bao gồm cả đáp án)
                            var blockText = string.Join(" ", block.SelectMany(el => el.Descendants<Run>().SelectMany(run => run.Elements<Text>().Select(t => t.Text))));
                            point = ExtractPointFromText(blockText);

                            // ➤ Tách phần `answerElements` ra từ block
                            var allParas = block.OfType<Paragraph>().ToList();
                            var firstAnswerPara = allParas.FirstOrDefault(p => Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));
                            if (firstAnswerPara != null)
                            {
                                answerElements = new List<OpenXmlElement> { firstAnswerPara };
                                int idx = block.IndexOf(firstAnswerPara);
                                for (int i = idx + 1; i < block.Count; i++)
                                {
                                    if (block[i] is Paragraph p && Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+")) break;
                                    answerElements.Add(block[i]);
                                }
                            }
                        }

                        foreach (var el in newBlock)
                            body.Append(el.CloneNode(true));

                        // ✅ Nếu có đáp án tự luận và đang có file đáp án mở
                        if (type == QuestionType.Essay && answerDoc != null)
                        {
                            var sourcePart = doc.MainDocumentPart!;
                            var targetPart = answerDoc.MainDocumentPart!;
                            var answerBody = targetPart.Document.Body ?? targetPart.Document.AppendChild(new Body());

                            // Chỉ tạo bảng nếu chưa có
                            var existingTable = answerBody.Elements<Table>().FirstOrDefault(t =>
                                t.Descendants<TableRow>().Any(r => r.InnerText.Contains("Câu") && r.InnerText.Contains("Đáp án")));

                            Table table;
                            if (existingTable != null)
                            {
                                table = existingTable;
                            }
                            else
                            {
                                table = CreateTable();
                                var header = new TableRow();
                                header.Append(CreateCell("Câu", "700"), CreateCell("Đáp án", "5000"), CreateCell("Điểm", "700"));
                                table.Append(header);
                                answerBody.Append(table);
                            }

                            // Tách phần đáp án tự luận
                            var extracted = await ExtractEssayAnswerAsync(block); // hàm bạn đã có hoặc tự viết

                            // Copy đáp án sang answerDoc
                            var contentClones = new List<OpenXmlElement>();
                            foreach (var el in extracted)
                            {
                                var clone = el.CloneNode(true);
                                // 👉 Nếu là Paragraph bắt đầu bằng "A. ", "B. ", ..., "Z. ", thì sửa ngay tại đây
                                if (clone is Paragraph p)
                                {
                                    // Gộp toàn bộ text trong Paragraph lại
                                    string combinedText = string.Concat(p.Descendants<Text>().Select(t => t.Text));

                                    // Nếu bắt đầu bằng "A. ", "B. ", ..., "Z. "
                                    if (Regex.IsMatch(combinedText, @"^[A-Z]\.\s+"))
                                    {
                                        // Xóa toàn bộ nội dung cũ
                                        p.RemoveAllChildren<Run>();

                                        // Xóa tiền tố "A. ", "B. ", ...
                                        string cleaned = Regex.Replace(combinedText, @"^[A-Z]\.\s+", "");

                                        // Tạo Run mới
                                        var run = new Run(new Text(cleaned));
                                        p.Append(run);
                                    }
                                }
                                foreach (var drawing in clone.Descendants<Drawing>())
                                {
                                    var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                                    if (blip?.Embed?.Value is string oldRelId)
                                    {
                                        if (sourcePart.GetPartById(oldRelId) is ImagePart sourceImage)
                                        {
                                            var newImagePart = targetPart.AddImagePart(sourceImage.ContentType);
                                            using var stream = sourceImage.GetStream();
                                            newImagePart.FeedData(stream);
                                            blip.Embed.Value = targetPart.GetIdOfPart(newImagePart);
                                        }
                                    }
                                }
                                contentClones.Add(clone);
                            }

                            // Trích điểm nếu có
                            var fullText = string.Join(" ", block.Select(b => b.InnerText));
                            point = ExtractPointFromText(fullText);

                            // Tạo hàng mới trong bảng
                            var row = new TableRow();
                            row.Append(CreateCell(localQuestion.ToString(), "700"), CreateCell(contentClones, "5000"), CreateCell(point ?? string.Empty, "700"));
                            table.Append(row);

                            // Đặt AnswerElements = null vì không dùng nữa
                            answerElements = null;
                        }

                        // Ghi vào danh sách câu hỏi
                        answers.Add(new QuestionExport
                        {
                            QuestionNumber = localQuestion,
                            CorrectAnswer = correct,
                            Type = type,
                            Point = point,
                            AnswerElements = answerElements
                        });
                    }
                }
                doc.MainDocumentPart.Document.Save();
                return answers;
            });
        }

        private List<OpenXmlElement> ShuffleAnswers(List<OpenXmlElement> block, QuestionType type, string version, MainDocumentPart sourcePart, out string correctAnswer, out List<OpenXmlElement>? answerElements)
        {
            var rnd = new Random();
            correctAnswer = string.Empty;
            var allElements = block;
            var allParas = allElements.OfType<Paragraph>().ToList();
            answerElements = null;

            if (type == QuestionType.MultipleChoice)
            {
                var answerStartParas = allParas
                    .Select((p, index) => new { Paragraph = p, Index = index, Text = p.InnerText.Trim() })
                    .Where(item => Regex.IsMatch(item.Text, @"^[A-D]\.")).ToList();

                if (answerStartParas.Count < 2) return block;

                var answerGroups = new List<List<OpenXmlElement>>();
                for (int i = 0; i < answerStartParas.Count; i++)
                {
                    var startPara = answerStartParas[i].Paragraph;
                    int startIndex = allElements.IndexOf(startPara);
                    int endIndex = (i < answerStartParas.Count - 1)
                        ? allElements.IndexOf(answerStartParas[i + 1].Paragraph)
                        : allElements.Count;
                    var group = allElements.Skip(startIndex).Take(endIndex - startIndex).ToList();
                    answerGroups.Add(group);
                }

                int firstAnswerIndexInBlock = allElements.IndexOf(answerStartParas.First().Paragraph);
                var questionElements = allElements.Take(firstAnswerIndexInBlock).ToList();

                var shuffled = version.Equals(Constants.ROOT_CODE) ? answerGroups : answerGroups.OrderBy(_ => rnd.Next()).ToList();
                var labels = new[] { "A.", "B.", "C.", "D." };

                correctAnswer = UpdateAnswerLabels(shuffled, labels);

                bool hasMultiElementAnswer = shuffled.Any(g => g.Count > 1);
                bool hasImageOrFormula = shuffled.Any(g => HasImageOrFormula(g.OfType<Paragraph>().ToList()));
                int maxLength = shuffled.Select(g => g.OfType<Paragraph>().FirstOrDefault()?.InnerText.Trim().Length ?? 0).DefaultIfEmpty().Max();

                int perLine = hasMultiElementAnswer ? 1 :
                              hasImageOrFormula ? 2 :
                              maxLength < 18 ? 4 :
                              maxLength < 36 ? 2 : 1;

                int[] tabPositions = perLine switch
                {
                    1 => new[] { Constants.TABSTOP_1 },
                    2 => new[] { Constants.TABSTOP_1, Constants.TABSTOP_3 },
                    _ => new[] { Constants.TABSTOP_1, Constants.TABSTOP_2, Constants.TABSTOP_3, Constants.TABSTOP_4 }
                };

                var resultParas = new List<OpenXmlElement>();

                if (perLine == 1)
                {
                    foreach (var group in shuffled)
                    {
                        var firstPara = group.OfType<Paragraph>().FirstOrDefault();
                        var para = new Paragraph
                        {
                            ParagraphProperties = new ParagraphProperties(new Tabs(
                                new TabStop() { Val = TabStopValues.Left, Position = Constants.TABSTOP_1 }))
                        };
                        para.Append(new Run(new TabChar()));

                        if (firstPara != null)
                        {
                            bool labelProcessed = false;
                            foreach (var node in firstPara.Elements())
                            {
                                var clonedNode = node.CloneNode(true);
                                if (clonedNode is Run run)
                                {
                                    var text = run.GetFirstChild<Text>();
                                    if (!labelProcessed && text != null && Regex.IsMatch(text.Text, @"^[A-D]\."))
                                    {
                                        run.RunProperties?.RemoveChild(run.RunProperties.Underline);
                                        labelProcessed = true;
                                    }
                                }
                                para.Append(clonedNode);
                            }
                            resultParas.Add(para);
                        }

                        foreach (var el in group.Skip(1))
                            resultParas.Add(el.CloneNode(true));
                    }
                }
                else
                {
                    for (int i = 0; i < shuffled.Count; i += perLine)
                    {
                        var lineGroups = shuffled.Skip(i).Take(perLine).ToList();

                        var para = new Paragraph
                        {
                            ParagraphProperties = new ParagraphProperties(new Tabs(
                                tabPositions.Select(tp => new TabStop() { Val = TabStopValues.Left, Position = tp })))
                        };
                        para.Append(new Run(new TabChar()));

                        for (int j = 0; j < lineGroups.Count; j++)
                        {
                            var group = lineGroups[j];
                            var firstPara = group.OfType<Paragraph>().FirstOrDefault();
                            if (firstPara != null)
                            {
                                bool labelProcessed = false;
                                foreach (var node in firstPara.Elements())
                                {
                                    var clonedNode = node.CloneNode(true);
                                    if (clonedNode is Run run)
                                    {
                                        var text = run.GetFirstChild<Text>();
                                        if (!labelProcessed && text != null && Regex.IsMatch(text.Text, @"^[A-D]\."))
                                        {
                                            run.RunProperties?.RemoveChild(run.RunProperties.Underline);
                                            labelProcessed = true;
                                        }
                                    }
                                    para.Append(clonedNode);
                                }
                            }
                            if (j < lineGroups.Count - 1)
                                para.Append(new Run(new TabChar()));
                        }

                        resultParas.Add(para);

                        int maxParaCount = lineGroups.Max(g => g.OfType<Paragraph>().Count());

                        for (int paraIndex = 1; paraIndex < maxParaCount; paraIndex++)
                        {
                            for (int j = 0; j < lineGroups.Count; j++)
                            {
                                var group = lineGroups[j];
                                var paraList = group.OfType<Paragraph>().ToList();
                                if (paraIndex < paraList.Count)
                                {
                                    var originalPara = paraList[paraIndex];
                                    var additionalPara = new Paragraph
                                    {
                                        ParagraphProperties = new ParagraphProperties(new Indentation
                                        {
                                            Left = tabPositions[j].ToString(),
                                            Hanging = "0"
                                        })
                                    };

                                    bool labelProcessed = false;
                                    foreach (var node in originalPara.Elements())
                                    {
                                        var clonedNode = node.CloneNode(true);
                                        if (clonedNode is Run run)
                                        {
                                            var text = run.GetFirstChild<Text>();
                                            if (!labelProcessed && text != null && Regex.IsMatch(text.Text, @"^[A-D]\."))
                                            {
                                                run.RunProperties?.RemoveChild(run.RunProperties.Underline);
                                                labelProcessed = true;
                                            }
                                        }
                                        additionalPara.Append(clonedNode);
                                    }

                                    resultParas.Add(additionalPara);
                                }
                            }
                        }

                        for (int j = 0; j < lineGroups.Count; j++)
                        {
                            var group = lineGroups[j];
                            var extraElements = group.Where(e => !(e is Paragraph)).Skip(1);
                            foreach (var el in extraElements)
                                resultParas.Add(el.CloneNode(true));
                        }
                    }
                }

                var result = new List<OpenXmlElement>();
                result.AddRange(questionElements.Select(e => e.CloneNode(true)));
                result.AddRange(resultParas);
                return result;
            }

            if (type == QuestionType.TrueFalse)
            {
                var trueFalseFirstParas = allParas
                    .Where(p => Regex.IsMatch(p.InnerText.Trim(), @"^[a-d]\)"))
                    .ToList();

                var answerGroups = new List<List<OpenXmlElement>>();

                for (int i = 0; i < trueFalseFirstParas.Count; i++)
                {
                    var startPara = trueFalseFirstParas[i];
                    int startIndex = allElements.IndexOf(startPara);
                    int endIndex = (i < trueFalseFirstParas.Count - 1)
                        ? allElements.IndexOf(trueFalseFirstParas[i + 1])
                        : allElements.Count;

                    answerGroups.Add(allElements.Skip(startIndex).Take(endIndex - startIndex).ToList());
                }

                var originalList = answerGroups.Select(group => new
                {
                    ParaGroup = group,
                    FirstPara = group.OfType<Paragraph>().FirstOrDefault(),
                    Label = group.OfType<Paragraph>().FirstOrDefault()?.InnerText.Trim().Substring(0, 2) ?? "",
                    IsCorrect = group.Any(el => el.Descendants<Run>().Any(r => r.RunProperties?.Underline?.Val != null &&
                                                                                r.RunProperties.Underline.Val != UnderlineValues.None))
                }).ToList();

                var shuffled = version.Equals(Constants.ROOT_CODE) ? originalList : originalList.OrderBy(_ => rnd.Next()).ToList();
                var labels = new[] { "a)", "b)", "c)", "d)" };

                var resultParas = new List<OpenXmlElement>();

                for (int i = 0; i < shuffled.Count && i < labels.Length; i++)
                {
                    var item = shuffled[i];
                    var newLabel = labels[i];
                    var group = item.ParaGroup;

                    var firstPara = group.OfType<Paragraph>().FirstOrDefault();
                    var newFirstPara = new Paragraph();

                    // Clone ParagraphProperties và Tabs
                    newFirstPara.ParagraphProperties = (ParagraphProperties?)firstPara?.ParagraphProperties?.CloneNode(true) ?? new ParagraphProperties();
                    newFirstPara.ParagraphProperties.RemoveAllChildren<Tabs>();
                    newFirstPara.ParagraphProperties.Append(new Tabs(new TabStop { Val = TabStopValues.Left, Position = Constants.TABSTOP_1 }));
                    newFirstPara.Append(new Run(new TabChar()));

                    // Thêm nhãn mới (in đậm)
                    var labelRun = new Run(new Text(newLabel));
                    labelRun.RunProperties = new RunProperties(new Bold());
                    newFirstPara.Append(labelRun);

                    bool labelProcessed = false;

                    foreach (var node in firstPara!.Elements())
                    {
                        if (!labelProcessed && node is Run run)
                        {
                            string runText = run.InnerText.TrimStart();
                            if (runText.StartsWith(item.Label))
                            {
                                string remaining = runText.Substring(item.Label.Length).TrimStart();
                                if (!string.IsNullOrEmpty(remaining))
                                {
                                    var newRun = new Run(new Text(remaining) { Space = SpaceProcessingModeValues.Preserve });
                                    newFirstPara.Append(newRun);
                                }
                                labelProcessed = true;
                                continue;
                            }
                        }

                        // Clone mọi thứ khác, giữ nguyên Equation, hình ảnh...
                        newFirstPara.Append(node.CloneNode(true));
                    }

                    resultParas.Add(newFirstPara);

                    // Thêm phần còn lại trong nhóm
                    foreach (var el in group.Skip(1))
                        resultParas.Add(el.CloneNode(true));
                }

                correctAnswer = string.Join(" ", shuffled.Select((item, i) => $"{labels[i]} {(item.IsCorrect ? "Đúng" : "Sai")}"));

                // Lấy các phần tử không thuộc bất kỳ nhóm đáp án nào
                var firstAnswerPara = trueFalseFirstParas.FirstOrDefault();
                var nonAnswerElements = (firstAnswerPara != null)
                    ? allElements.Take(allElements.IndexOf(firstAnswerPara)).ToList()
                    : allElements.ToList();

                var result = new List<OpenXmlElement>();
                result.AddRange(nonAnswerElements.Select(e => (OpenXmlElement)e.CloneNode(true)));
                result.AddRange(resultParas);

                return result;
            }

            if (type == QuestionType.ShortAnswer)
            {
                var para = allElements.OfType<Paragraph>().FirstOrDefault(p => Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));

                if (para != null)
                {
                    // Trích nội dung đúng từ toàn bộ node (kể cả công thức, hình ảnh)
                    correctAnswer = "";

                    foreach (var node in para.Elements())
                    {
                        if (node is Run run)
                        {
                            var text = run.GetFirstChild<Text>();
                            if (text != null && Regex.IsMatch(text.Text, @"^[A-Z]\.\s*"))
                            {
                                var trimmed = Regex.Replace(text.Text, @"^[A-Z]\.\s*", "");
                                correctAnswer += trimmed;
                            }
                            else
                            {
                                correctAnswer += run.InnerText;
                            }
                        }
                        else
                        {
                            correctAnswer += node.InnerText;
                        }
                    }

                    correctAnswer = correctAnswer.Trim();

                    // Trả lại tất cả element trừ phần chứa đáp án
                    return allElements.Where(e => e != para)
                                      .Select(e => (OpenXmlElement)e.CloneNode(true))
                                      .ToList();
                }
                else
                {
                    correctAnswer = string.Empty;
                    return allElements.Select(x => (OpenXmlElement)x.CloneNode(true)).ToList();
                }
            }

            if (type == QuestionType.Essay)
            {
                var firstAnswerPara = allParas.FirstOrDefault(p => Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));
                if (firstAnswerPara != null)
                {
                    int firstAnswerIndex = allParas.IndexOf(firstAnswerPara);
                    int firstAnswerIndexInBlock = allElements.IndexOf(firstAnswerPara);

                    var extractedAnswerElements = new List<OpenXmlElement> { firstAnswerPara };

                    // Tìm các phần tử tiếp theo thuộc phần đáp án
                    for (int i = allElements.IndexOf(firstAnswerPara) + 1; i < allElements.Count; i++)
                    {
                        if (allElements[i] is Paragraph p && Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"))
                            break;

                        extractedAnswerElements.Add(allElements[i]);
                    }

                    // Sao chép & sửa hình ảnh
                    var clonedAnswerElements = new List<OpenXmlElement>();
                    foreach (var el in extractedAnswerElements)
                    {
                        var cloned = el.CloneNode(true);

                        foreach (var drawing in cloned.Descendants<Drawing>())
                        {
                            var blip = drawing.Descendants<DocumentFormat.OpenXml.Drawing.Blip>().FirstOrDefault();
                            if (blip?.Embed?.Value is string oldRelId)
                            {
                                if (sourcePart.GetPartById(oldRelId) is ImagePart sourceImage)
                                {
                                    var newImagePart = sourcePart.AddImagePart(sourceImage.ContentType);
                                    using var stream = sourceImage.GetStream();
                                    newImagePart.FeedData(stream);

                                    var newRelId = sourcePart.GetIdOfPart(newImagePart);
                                    blip.Embed.Value = newRelId;
                                }
                            }
                        }

                        clonedAnswerElements.Add(cloned);
                    }

                    answerElements = clonedAnswerElements;

                    // Gộp nội dung từ các phần tử trong đáp án
                    var builder = new StringBuilder();

                    foreach (var el in answerElements)
                    {
                        builder.AppendLine(el.InnerText.Trim());
                    }

                    correctAnswer = builder.ToString().Trim();

                    // Trả lại phần đề (không gồm đáp án)
                    var result = allElements
                        .Where(el => !extractedAnswerElements.Contains(el))
                        .Select(el => (OpenXmlElement)el.CloneNode(true))
                        .ToList();

                    return result;
                }
                else
                {
                    correctAnswer = string.Empty;
                    answerElements = null;
                    return block.Select(e => (OpenXmlElement)e.CloneNode(true)).ToList();
                }
            }

            correctAnswer = string.Empty;
            answerElements = null;
            return block;
        }

        private QuestionType DetectQuestionType(List<OpenXmlElement> block)
        {
            var paras = block.OfType<Paragraph>().ToList();
            var answers = paras.Where(p => Regex.IsMatch(p.InnerText.Trim(), @"^[A-D]\.")).ToList();
            var trueFalse = paras.Where(p => Regex.IsMatch(p.InnerText.Trim(), @"^[a-d]\)\s")).ToList();

            if (answers.Count >= 2) return QuestionType.MultipleChoice;
            if (trueFalse.Count >= 2) return QuestionType.TrueFalse;
            if (answers.Count == 1)
            {
                var content = answers[0].InnerText.Substring(2).Trim();
                return content.Length > 8 ? QuestionType.Essay : QuestionType.ShortAnswer;
            }

            return QuestionType.ShortAnswer;
        }

        private List<List<OpenXmlElement>> SplitQuestions(Body body)
        {
            var result = new List<List<OpenXmlElement>>();
            var current = new List<OpenXmlElement>();
            Regex headerRegex = Constants.QuestionHeaderRegex;

            foreach (var el in body.Elements())
            {
                if (el is Paragraph para)
                {
                    var text = para.InnerText.Trim();
                    if (headerRegex.IsMatch(text))
                    {
                        if (current.Count > 0)
                            result.Add(current);
                        current = new List<OpenXmlElement>();
                    }
                }
                current.Add(el.CloneNode(true));
            }

            if (current.Count > 0)
                result.Add(current);

            return result;
        }

        public async Task<List<Question>> ParseDocxQuestionsAsync(string filePath)
        {
            return await Task.Run(() =>
            {
                using var doc = WordprocessingDocument.Open(filePath, false);
                var paragraphs = doc.MainDocumentPart!.Document.Body!.Elements<Paragraph>().ToList();
                var questions = new List<Question>();

                Regex questionHeader = Constants.QuestionHeaderRegex;
                Regex mcAnswerRegex = Constants.MultipleChoiceAnswerRegex;
                Regex trueFalseAnswer = Constants.TrueFalseAnswerRegex;
                Regex levelRegex = Constants.LevelRegex;

                int code = 1;
                for (int i = 0; i < paragraphs.Count; i++)
                {
                    string text = paragraphs[i].InnerText.Trim();
                    if (!questionHeader.IsMatch(text)) continue;

                    var question = new Question
                    {
                        Code = code++,
                        Level = Level.Know,
                        IsValid = false // mặc định
                    };

                    // Lấy mức độ từ tiêu đề nếu có
                    var levelMatch = levelRegex.Match(text);
                    if (levelMatch.Success)
                    {
                        question.Level = GetLevelFromText(levelMatch.Value);
                        text = levelRegex.Replace(text, "").Trim(); // loại bỏ ký hiệu (NB) khỏi nội dung câu
                    }

                    // ✅ Lấy nội dung câu hỏi
                    var questionTextRuns = new List<Run>();
                    questionTextRuns.AddRange(paragraphs[i].Elements<Run>());
                    int j = i + 1;

                    while (j < paragraphs.Count && !questionHeader.IsMatch(paragraphs[j].InnerText.Trim()))
                    {
                        var para = paragraphs[j];
                        string line = para.InnerText.Trim();

                        // Khi gặp dòng bắt đầu là đáp án thì dừng thu thập nội dung câu hỏi
                        if (mcAnswerRegex.IsMatch(line) || trueFalseAnswer.IsMatch(line))
                            break;

                        questionTextRuns.AddRange(para.Elements<Run>());
                        j++;
                    }

                    // Ghép nội dung từ tất cả các run lại (bao gồm cả dòng tiêu đề)
                    string rawText = string.Join("", questionTextRuns.Select(r => r.InnerText).Where(t => !string.IsNullOrWhiteSpace(t))).Trim();

                    // Tách phần "Câu X." ra nếu có, rồi thêm dấu cách
                    var match = Regex.Match(rawText, @"^(Câu\s+\d+[\.:]?)\s*(.*)", RegexOptions.IgnoreCase);
                    question.QuestionText = match.Success ? $"{match.Groups[1].Value} {match.Groups[2].Value.Trim()}" : rawText;

                    // ✅ Lấy danh sách đáp án và đáp án đúng
                    var answers = new List<string>();
                    var correctAnswers = new List<string>();

                    while (j < paragraphs.Count && !questionHeader.IsMatch(paragraphs[j].InnerText.Trim()))
                    {
                        var para = paragraphs[j];
                        string line = para.InnerText.Trim();

                        if (mcAnswerRegex.IsMatch(line))
                        {
                            answers.Add(line);
                            string label = line.Substring(0, 2); // ví dụ "A.", "a)", ...
                            if (IsUnderlined(para, label))
                                correctAnswers.Add(label.TrimEnd('.', ')'));
                        }
                        else if (trueFalseAnswer.IsMatch(line))
                        {
                            string label = line.Substring(0, 2);
                            string indicator = IsUnderlined(para, label) ? "Đ" : "S";
                            var levelMatchItem = levelRegex.Match(line);
                            var level = levelMatchItem.Success ? GetLevelFromText(levelMatchItem.Value) : Level.Know;

                            string formatted = $"{label} {indicator} ({ShortLevelCode(level)})";
                            answers.Add(line);
                            correctAnswers.Add(formatted);
                        }

                        j++;
                    }

                    question.CountAnswer = answers.Count;
                    question.Answers = string.Join("\n", answers);
                    question.CorrectAnswer = string.Join(" | ", correctAnswers);

                    // ✅ Phân loại câu hỏi và đánh giá IsValid
                    if (answers.Count > 1 && answers.All(a => mcAnswerRegex.IsMatch(a)))
                    {
                        question.QuestionType = QuestionType.MultipleChoice;

                        bool validCount = answers.Count == 4;
                        bool oneCorrect = correctAnswers.Count == 1;

                        if (!validCount)
                            question.Description += answers.Count < 4 ? "⚠️ Không đủ 4 đáp án. " : "⚠️ Vượt quá 4 đáp án. ";
                        if (!oneCorrect)
                            question.Description += "⚠️ Phải có đúng 1 đáp án đúng. ";

                        question.IsValid = validCount && oneCorrect;
                        if (question.IsValid) question.Description += "✅ OK";
                    }
                    else if (answers.Count > 1 && answers.All(a => trueFalseAnswer.IsMatch(a)))
                    {
                        question.QuestionType = QuestionType.TrueFalse;
                        question.Level = Level.None;
                        question.IsValid = answers.Count == 4;

                        if (!question.IsValid)
                            question.Description += "⚠️ Câu đúng/sai phải có đúng 4 ý a), b), c), d). ";
                        else
                            question.Description += "✅ OK";
                    }
                    else if (answers.Count == 1 && mcAnswerRegex.IsMatch(answers[0]))
                    {
                        string content = answers[0].Substring(2).Trim();
                        question.CorrectAnswer = content;

                        if (string.IsNullOrWhiteSpace(content))
                        {
                            question.QuestionType = QuestionType.ShortAnswer;
                            question.Description += "⚠️ Chưa có nội dung đáp án. ";
                            question.IsValid = false;
                        }
                        else if (content.Length <= 4)
                        {
                            question.QuestionType = QuestionType.ShortAnswer;
                            question.IsValid = true;
                            question.Description += "✅ OK";
                        }
                        else
                        {
                            question.QuestionType = QuestionType.Essay;

                            var extractedAnswers = new List<OpenXmlElement>();
                            while (j < paragraphs.Count && !questionHeader.IsMatch(paragraphs[j].InnerText.Trim()))
                            {
                                extractedAnswers.Add(paragraphs[j]);
                                j++;
                            }

                            // Kiểm tra danh sách extractedAnswers có phần tử nào không
                            if (extractedAnswers.Any())
                            {
                                question.CorrectAnswer = string.Join("\n", extractedAnswers.Select(e => e.InnerText.Trim()));

                                // Đảm bảo rằng dữ liệu không bị rỗng
                                if (string.IsNullOrWhiteSpace(question.CorrectAnswer))
                                {
                                    question.CorrectAnswer = "⚠️ Nội dung trống, kiểm tra lại cách lấy dữ liệu!";
                                }
                            }
                            else
                            {
                                question.CorrectAnswer = "⚠️ Không tìm thấy nội dung đáp án!";
                            }

                            question.IsValid = true;
                            question.Description += "✅ OK";
                        }
                    }
                    else if (answers.Count == 1)
                    {
                        question.QuestionType = QuestionType.Unknown;
                        question.Description += "⚠️ Đáp án không đúng định dạng bắt đầu bằng 'A.', 'B.', ... ";
                        question.IsValid = false;
                    }
                    else
                    {
                        question.QuestionType = QuestionType.Unknown;
                        question.Description += "⚠️ Không nhận dạng được dạng câu.";
                        question.IsValid = false;
                    }

                    questions.Add(question);
                    i = j - 1;
                }

                return questions;
            });
        }

        private async Task AddEndNotesAsync(WordprocessingDocument doc)
        {
            await Task.Run(() =>
            {
                var body = doc.MainDocumentPart!.Document.Body;

                // Tạo đoạn văn bản đầu tiên (HẾT) in đậm và căn giữa
                var endNote1 = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }),
                                              new Run(new Text("-------------------- HẾT --------------------") { Space = SpaceProcessingModeValues.Preserve })
                                              {
                                                  RunProperties = new RunProperties(new Bold())
                                              });

                // Tạo đoạn văn bản thứ hai (Thí sinh không được sử dụng tài liệu) in nghiêng
                var endNote2 = new Paragraph(new Run(new Text("- Thí sinh không được sử dụng tài liệu;") { Space = SpaceProcessingModeValues.Preserve })
                {
                    RunProperties = new RunProperties(new Italic())
                });

                // Tạo đoạn văn bản thứ ba (Giám thị không giải thích gì thêm) in nghiêng
                var endNote3 = new Paragraph(new Run(new Text("- Giám thị không giải thích gì thêm.") { Space = SpaceProcessingModeValues.Preserve })
                {
                    RunProperties = new RunProperties(new Italic())
                });

                // Thêm các đoạn vào cuối body
                body!.Append(endNote1);
                body.Append(endNote2);
                body.Append(endNote3);

                doc.MainDocumentPart.Document.Save();
            });
        }

        public async Task AddFooterAsync(WordprocessingDocument doc, string version)
        {
            await Task.Run(() =>
            {
                var footerPart = doc.MainDocumentPart!.AddNewPart<FooterPart>();
                var footer = new Footer();

                // Tạo phần thông tin trang
                var paragraph = new Paragraph();

                // Căn phải cho đoạn văn
                var paragraphProperties = new ParagraphProperties();
                paragraphProperties.Append(new Justification() { Val = JustificationValues.Right });
                paragraph.AppendChild(paragraphProperties);

                paragraph.Append(new Run(new Text($"Trang ") { Space = SpaceProcessingModeValues.Preserve }));

                // Trường cho số trang hiện tại
                var pageNumberField = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                var pageNumber = new Run(new FieldCode("PAGE"));
                var pageNumberEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });

                // Trường cho tổng số trang
                var totalPagesField = new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin });
                var totalPages = new Run(new FieldCode("SECTIONPAGES"));
                var totalPagesEnd = new Run(new FieldChar() { FieldCharType = FieldCharValues.End });

                // Thêm thông tin vào paragraph
                paragraph.Append(pageNumberField, pageNumber, pageNumberEnd);
                paragraph.Append(new Run(new Text($"/")));
                paragraph.Append(totalPagesField, totalPages, totalPagesEnd);
                paragraph.Append(new Run(new Text($" - Mã đề {version}") { Space = SpaceProcessingModeValues.Preserve }));

                footer.Append(paragraph);
                footerPart.Footer = footer;

                // Thêm footer vào tài liệu
                var sectionProperties = doc.MainDocumentPart.Document.Body!.Elements<SectionProperties>().FirstOrDefault();
                if (sectionProperties != null)
                {
                    var footerReference = new FooterReference() { Id = doc.MainDocumentPart.GetIdOfPart(footerPart), Type = HeaderFooterValues.Default };
                    sectionProperties.Append(footerReference);
                }
                doc.MainDocumentPart.Document.Save();
            });
        }

        private async Task InsertTemplateAsync(string templatePath, WordprocessingDocument doc, MixInfo mixInfo, string code)
        {
            await Task.Run(async () =>
            {
                var targetBody = doc.MainDocumentPart!.Document.Body!;

                using var templateDoc = WordprocessingDocument.Open(templatePath, false);
                var templateBody = templateDoc.MainDocumentPart!.Document.Body!;

                foreach (var element in templateBody.Elements().Reverse())
                {
                    targetBody.InsertAt(element.CloneNode(true), 0);
                }

                await ReplacePlaceholdersAsync(targetBody, new Dictionary<string, string>
                {
                    { "[KYTHI]", mixInfo.TestPeriod ?? string.Empty },
                    { "[NAMHOC]", mixInfo.SchoolYear ?? string.Empty },
                    { "[MONTHI]", mixInfo.Subject ?? string.Empty },
                    { "[DONVICAPTREN]", mixInfo.SuperiorUnit ?? string.Empty },
                    { "[DONVI]", mixInfo.Unit ?? string.Empty },
                    { "[MaDe]", code },
                    { "[ThoiGian]", mixInfo.Time ?? string.Empty }
                });

                doc.MainDocumentPart.Document.Save();
            });
        }

        private async Task ReplacePlaceholdersAsync(Body body, Dictionary<string, string> replacements)
        {
            await Task.Run(() =>
            {
                foreach (var para in body.Descendants<Paragraph>())
                {
                    UpdateTextPlaceholders(para, replacements);
                    InsertNumPagesField(para);
                }
            });
        }

        private async Task AppendGuideAsync(WordprocessingDocument doc, List<QuestionExport> answers, MixInfo mixInfo, string code)
        {
            await Task.Run(async () =>
            {
                try
                {
                    var mainPart = doc.MainDocumentPart!;
                    var document = mainPart.Document;
                    var body = document.Body ?? document.AppendChild(new Body());

                    await ReplacePlaceholdersAsync(body!, new Dictionary<string, string>
                    {
                        { "[KYTHI]", mixInfo.TestPeriod ?? string.Empty },
                        { "[NAMHOC]", mixInfo.SchoolYear ?? string.Empty },
                        { "[MONTHI]", mixInfo.Subject ?? string.Empty },
                        { "[DONVICAPTREN]", mixInfo.SuperiorUnit ?? string.Empty },
                        { "[DONVI]", mixInfo.Unit ?? string.Empty },
                        { "[MaDe]", code },
                        { "[ThoiGian]", mixInfo.Time ?? string.Empty }
                    });

                    var grouped = answers.GroupBy(a => a.Type).OrderBy(g => g.Key).ToList();

                    int index = 0;
                    foreach (var group in grouped)
                    {
                        index++;
                        string title = CreateSectionTitle(group.Key, index, group.Count());

                        if (!string.IsNullOrEmpty(title))
                        {
                            var heading = new Paragraph();
                            var parts = title.Split('.', 3);
                            if (parts.Length >= 3)
                            {
                                var boldRun = new Run(new RunProperties(new Bold()), new Text($"{parts[0]}.{parts[1]}.") { Space = SpaceProcessingModeValues.Preserve });
                                var normalRun = new Run(new Text(parts[2]) { Space = SpaceProcessingModeValues.Preserve });
                                heading.Append(boldRun, normalRun);
                            }
                            body.Append(heading);
                        }

                        var table = CreateTable();

                        if (group.Key == QuestionType.MultipleChoice)
                        {
                            body.Append(new Paragraph(new Run(new Text("Mỗi câu trả lời đúng thí sinh được 0,25 điểm.") { Space = SpaceProcessingModeValues.Preserve })));
                            AddMultipleChoiceAnswerTable(table, answers.Where(q => q.Type == QuestionType.MultipleChoice).ToList());
                            body.Append(table);
                        }
                        else if (group.Key == QuestionType.TrueFalse)
                        {
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác 01 ý trong 01 câu hỏi được 0,1 điểm;") { Space = SpaceProcessingModeValues.Preserve })));
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác 02 ý trong 01 câu hỏi được 0,25 điểm;") { Space = SpaceProcessingModeValues.Preserve })));
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác 03 ý trong 01 câu hỏi được 0,5 điểm;") { Space = SpaceProcessingModeValues.Preserve })));
                            body.Append(new Paragraph(new Run(new Text("- Thí sinh chỉ lựa chọn chính xác cả 04 ý trong 01 câu hỏi được 1 điểm.") { Space = SpaceProcessingModeValues.Preserve })));
                            AddTrueFalseAnswerTable(table, group.Select(a => a.CorrectAnswer).ToList()!, 8);
                            body.Append(table);
                        }
                        else if (group.Key == QuestionType.ShortAnswer)
                        {
                            body.Append(new Paragraph(new Run(new Text($"Mỗi câu trả lời đúng thí sinh được 0,5 điểm.") { Space = SpaceProcessingModeValues.Preserve })));
                            AddShortAnswerTable(table, group.Select(a => a.CorrectAnswer).ToList()!, 6);
                            body.Append(table);
                        }
                        else if (group.Key == QuestionType.Essay)
                        {
                            // TODO
                        }

                        mainPart.Document.Save();
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error("DOCX save error: " + ex.Message);
                }
            });
        }

        private Table CreateTable()
        {
            return new Table(
                new TableProperties(
                    new TableWidth { Width = "5000", Type = TableWidthUnitValues.Pct },
                    new TableLayout { Type = TableLayoutValues.Autofit },
                    new TableJustification { Val = TableRowAlignmentValues.Center },
                    new TableBorders(
                        new TopBorder { Val = BorderValues.Single, Size = 4 },
                        new BottomBorder { Val = BorderValues.Single, Size = 4 },
                        new LeftBorder { Val = BorderValues.Single, Size = 4 },
                        new RightBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideHorizontalBorder { Val = BorderValues.Single, Size = 4 },
                        new InsideVerticalBorder { Val = BorderValues.Single, Size = 4 }
                    )
                )
            );
        }

        private TableCell CreateCell(string text, string width = "1200")
        {
            var lines = text.Split('\n');
            bool isMultiLine = lines.Length > 1;

            var cell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                )
            );

            var paragraph = new Paragraph(
                new ParagraphProperties(
                    new Justification() { Val = isMultiLine ? JustificationValues.Left : JustificationValues.Center }
                )
            );

            for (int i = 0; i < lines.Length; i++)
            {
                paragraph.Append(new Run(new Text(lines[i]) { Space = SpaceProcessingModeValues.Preserve }));
                if (i < lines.Length - 1)
                {
                    paragraph.Append(new Break());
                }
            }
            cell.Append(paragraph);
            return cell;
        }

        private TableCell CreateCell(List<OpenXmlElement> elements, string width = "5000")
        {
            var cell = new TableCell(
                new TableCellProperties(
                    new TableCellWidth { Type = TableWidthUnitValues.Dxa, Width = width },
                    new TableCellVerticalAlignment { Val = TableVerticalAlignmentValues.Center }
                )
            );

            foreach (var element in elements)
            {
                if (element is Paragraph || element is Table || element is Drawing || element is DocumentFormat.OpenXml.Math.OfficeMath)
                {
                    var cloned = element.CloneNode(true);
                    cell.Append(cloned);
                }
                // Nếu là run đơn lẻ hoặc đoạn trống thì có thể đóng gói lại
                else if (element is Run run)
                {
                    var para = new Paragraph();
                    para.Append(run.CloneNode(true));
                    cell.Append(para);
                }
            }

            return cell;
        }

        private void AddMultipleChoiceAnswerTable(Table table, List<QuestionExport> mcQuestions)
        {
            // ==== 2. Tính số lượng câu và làm tròn lên mốc chẵn 10 ====
            int count = mcQuestions.Count;
            int roundedCount = ((count + 9) / 10) * 10;

            // ==== 3. Tạo từng nhóm 10 câu ====
            for (int i = 0; i < roundedCount; i += 10)
            {
                var numberRow = new TableRow();
                var answerRow = new TableRow();

                // Cột đầu tiên: "Câu" và "Đáp án"
                numberRow.Append(CreateCell("Câu"));
                answerRow.Append(CreateCell("Đáp án"));

                for (int j = i + 1; j <= i + 10; j++)
                {
                    numberRow.Append(CreateCell(j <= count ? j.ToString() : string.Empty));
                    var ans = mcQuestions.FirstOrDefault(q => q.QuestionNumber == j);
                    answerRow.Append(CreateCell(ans?.CorrectAnswer ?? string.Empty));
                }

                table.Append(numberRow);
                table.Append(answerRow);
            }
        }

        private void AddTrueFalseAnswerTable(Table table, List<string> answers, int maxColumns)
        {
            // Hàng 1: Câu 1 → N
            var headerRow = new TableRow();
            headerRow.Append(CreateCell("Câu"));
            for (int i = 1; i <= maxColumns; i++)
            {
                headerRow.Append(CreateCell(i.ToString()));
            }
            table.Append(headerRow);

            // Hàng 2: Đáp án (gộp a-d vào 1 ô, mỗi dòng là một lựa chọn)
            var answerRow = new TableRow();
            answerRow.Append(CreateCell("Đáp án"));

            for (int i = 0; i < maxColumns; i++)
            {
                if (i < answers.Count)
                {
                    var lines = answers[i]
                        .Split(new[] { ' ' }, StringSplitOptions.RemoveEmptyEntries)
                        .Chunk(2)
                        .Select(p => $"{p[0]} {p[1]}")
                        .ToList();

                    string combined = string.Join(Environment.NewLine, lines);
                    answerRow.Append(CreateCell(combined));
                }
                else
                {
                    answerRow.Append(CreateCell(""));
                }
            }
            table.Append(answerRow);
        }

        private void AddShortAnswerTable(Table table, List<string> answers, int maxColumns)
        {
            // Hàng 1: Câu 1 → N
            var headerRow = new TableRow();
            headerRow.Append(CreateCell("Câu"));
            for (int i = 1; i <= maxColumns; i++)
            {
                headerRow.Append(CreateCell(i.ToString()));
            }
            table.Append(headerRow);

            // Hàng 2: Đáp án (gộp a-d vào 1 ô, mỗi dòng là một lựa chọn)
            var answerRow = new TableRow();
            answerRow.Append(CreateCell("Đáp án"));
            for (int i = 0; i < maxColumns; i++)
            {
                answerRow.Append(CreateCell(i < answers.Count ? answers[i] : ""));
            }
            table.Append(answerRow);
        }

        private bool IsUnderlined(Paragraph para, string label)
        {
            foreach (var run in para.Elements<Run>())
            {
                string runText = run.InnerText.Trim();
                if (runText.StartsWith(label, StringComparison.OrdinalIgnoreCase))
                {
                    var underline = run.RunProperties?.Underline;
                    return underline != null && underline.Val != null && underline.Val != UnderlineValues.None;
                }
            }

            return false;
        }

        private Level GetLevelFromText(string value)
        {
            if (value.Contains("TH", StringComparison.OrdinalIgnoreCase)) return Level.Understand;
            if (value.Contains("VD", StringComparison.OrdinalIgnoreCase)) return Level.Manipulate;
            return Level.Know;
        }

        private string ShortLevelCode(Level level) => level switch
        {
            Level.Know => "NB",
            Level.Understand => "TH",
            Level.Manipulate => "VD",
            _ => ""
        };

        private void FormatParagraph(Paragraph para)
        {
            para.ParagraphProperties ??= new ParagraphProperties();
            para.ParagraphProperties.SpacingBetweenLines = new SpacingBetweenLines
            {
                Before = "0",
                After = "0",
                Line = "288", // 1.2 dòng = 240 * 1.2 = 288 twips
                LineRule = LineSpacingRuleValues.Auto
            };

            foreach (var run in para.Elements<Run>())
            {
                run.RunProperties ??= new RunProperties();
                run.RunProperties.RunFonts = new RunFonts { Ascii = "Times New Roman", HighAnsi = "Times New Roman" };
                run.RunProperties.FontSize = new FontSize { Val = "24" }; // 12pt
            }
        }

        private bool HasImageOrFormula(List<Paragraph> answerGroup)
        {
            return answerGroup.Any(para =>
                para.Descendants<Drawing>().Any() ||                                // Hình ảnh (Drawing)
                para.Descendants<DocumentFormat.OpenXml.Math.OfficeMath>().Any() || // Công thức toán
                para.Descendants<EmbeddedObject>().Any() ||                         // Object nhúng
                para.Descendants<DocumentFormat.OpenXml.Vml.Shape>().Any() ||       // VML Shape
                para.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().Any()      // VML Image (format cũ)
            );
        }

        private string CreateSectionTitle(QuestionType type, int index, int end)
        {
            return type switch
            {
                QuestionType.MultipleChoice => string.Format(Constants.TITLES[0], Constants.ROMANS[index - 1], end),
                QuestionType.TrueFalse => string.Format(Constants.TITLES[1], Constants.ROMANS[index - 1], end),
                QuestionType.ShortAnswer => string.Format(Constants.TITLES[2], Constants.ROMANS[index - 1], end),
                QuestionType.Essay => string.Format(Constants.TITLES[3], Constants.ROMANS[index - 1], end),
                _ => string.Empty
            };
        }

        public string? ExtractPointFromText(string text) => Regex.Match(text, @"\((\d+[,.]?\d*)\s+điểm\)").Groups[1].Value;

        private void UpdateTextPlaceholders(Paragraph para, Dictionary<string, string> replacements)
        {
            var runs = para.Elements<Run>().ToList();
            if (!runs.Any()) return;

            var fullText = string.Join("", runs.Select(r => string.Concat(r.Elements<Text>().Select(t => t.Text ?? ""))));
            if (!replacements.Keys.Any(k => fullText.Contains(k))) return;

            string modifiedText = fullText;
            foreach (var kvp in replacements)
                modifiedText = modifiedText.Replace(kvp.Key, kvp.Value);

            para.RemoveAllChildren<Run>();
            var newRun = new Run(new Text(modifiedText) { Space = SpaceProcessingModeValues.Preserve });

            if (runs.FirstOrDefault()?.RunProperties != null)
                newRun.RunProperties = (RunProperties)runs.First().RunProperties!.CloneNode(true);

            para.AppendChild(newRun);
        }

        private void InsertNumPagesField(Paragraph para)
        {
            var runs = para.Elements<Run>().ToList();
            if (!runs.Any()) return;

            var fullText = string.Join("", runs.Select(r => string.Concat(r.Elements<Text>().Select(t => t.Text ?? ""))));
            if (!fullText.Contains("[NUMPAGES]")) return;

            RunProperties? originalProps = runs.FirstOrDefault()?.RunProperties?.CloneNode(true) as RunProperties;
            string[] parts = fullText.Split(new[] { "[NUMPAGES]" }, StringSplitOptions.None);

            para.RemoveAllChildren<Run>();

            for (int i = 0; i < parts.Length; i++)
            {
                if (!string.IsNullOrEmpty(parts[i]))
                {
                    var run = new Run(new Text(parts[i]) { Space = SpaceProcessingModeValues.Preserve });
                    if (originalProps != null) run.RunProperties = (RunProperties)originalProps.CloneNode(true);
                    para.AppendChild(run);
                }

                if (i < parts.Length - 1)
                {
                    para.Append(
                        CreateFieldRun(FieldCharValues.Begin, originalProps),
                        CreateCodeRun(" SECTIONPAGES ", originalProps),
                        CreateFieldRun(FieldCharValues.Separate, originalProps),
                        CreateResultRun("1", originalProps),
                        CreateFieldRun(FieldCharValues.End, originalProps)
                    );
                }
            }
        }

        private Run CreateFieldRun(FieldCharValues type, RunProperties? props)
        {
            var run = new Run { RunProperties = props?.CloneNode(true) as RunProperties };
            run.AppendChild(new FieldChar { FieldCharType = type });
            return run;
        }

        private Run CreateCodeRun(string code, RunProperties? props) =>
            new Run(new FieldCode(code) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props?.CloneNode(true) as RunProperties };

        private Run CreateResultRun(string result, RunProperties? props) =>
            new Run(new Text(result) { Space = SpaceProcessingModeValues.Preserve }) { RunProperties = props?.CloneNode(true) as RunProperties };

        private async Task UpdateQuestionNumberAsync(Paragraph para, int localQuestion)
        {
            await Task.Run(() =>
            {
                var text = para.Elements<Run>().FirstOrDefault()?.Elements<Text>().FirstOrDefault();
                if (text != null)
                    text.Text = Regex.Replace(text.Text, @"^Câu\s+\d+", $"Câu {localQuestion}");
            });
        }

        private string UpdateAnswerLabels(List<List<OpenXmlElement>> shuffled, string[] labels)
        {
            string correctAnswer = string.Empty;

            for (int i = 0; i < shuffled.Count && i < labels.Length; i++)
            {
                var group = shuffled[i];
                bool isCorrect = group.Any(el =>
                    el.Descendants<Run>().Any(r => r.RunProperties?.Underline?.Val != null &&
                                                   r.RunProperties.Underline.Val != UnderlineValues.None));
                if (isCorrect)
                    correctAnswer = labels[i].Substring(0, 1);

                var firstPara = group.OfType<Paragraph>().FirstOrDefault();
                if (firstPara != null)
                {
                    var firstTextRun = firstPara.Elements<Run>().FirstOrDefault();
                    var firstText = firstTextRun?.GetFirstChild<Text>();
                    if (firstText != null)
                        firstText.Text = Regex.Replace(firstText.Text, @"^[A-D]\.", labels[i]);
                }
            }
            return correctAnswer;
        }

        private async Task<List<OpenXmlElement>> ExtractEssayAnswerAsync(List<OpenXmlElement> block)
        {
            return await Task.Run(() =>
            {
                var answerElements = new List<OpenXmlElement>();

                // Tìm paragraph đầu tiên có dạng "A. ..."
                var paras = block.OfType<Paragraph>().ToList();
                var firstAnswerPara = paras.FirstOrDefault(p => Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"));

                if (firstAnswerPara == null)
                    return answerElements;

                int startIndex = block.IndexOf(firstAnswerPara);

                // Thêm phần tử đầu tiên
                answerElements.Add(block[startIndex]);

                // Thêm các phần tử tiếp theo cho đến khi gặp một đoạn có dạng "A. ..." khác (hiếm khi có)
                for (int i = startIndex + 1; i < block.Count; i++)
                {
                    if (block[i] is Paragraph p && Regex.IsMatch(p.InnerText.Trim(), @"^[A-Z]\.\s+"))
                        break; // gặp nhãn mới → dừng
                    answerElements.Add(block[i]);
                }

                return answerElements;
            });
        }

        private async Task MoveEssayTableToEndAsync(WordprocessingDocument answerDoc)
        {
            await Task.Run(() =>
            {
                var body = answerDoc.MainDocumentPart?.Document.Body;
                if (body == null) return;

                var tables = body.Elements<Table>().ToList();

                var essayTable = tables.FirstOrDefault(t =>
                    t.Descendants<TableRow>().Any(r => r.InnerText.Contains("Đáp án")) &&
                    t.InnerText.Contains("Câu") &&
                    t.InnerText.Contains("Điểm"));

                if (essayTable != null)
                {
                    // Clone đúng cách bằng XML để giữ nguyên toàn bộ TableProperties
                    string xml = essayTable.OuterXml;
                    var newTable = new Table(xml);

                    essayTable.Remove(); // Xóa bản gốc
                    body.Append(newTable); // Thêm lại bản sao đúng
                }
            });
        }

        private void ProcessVmlElements(OpenXmlElement element, MainDocumentPart sourceMainPart, MainDocumentPart targetMainPart)
        {
            // Xử lý tất cả VML shapes, không chỉ những cái có ImageData
            foreach (var vmlShape in element.Descendants<DocumentFormat.OpenXml.Vml.Shape>())
            {
                // Xử lý ImageData nếu có
                var imageData = vmlShape.Descendants<DocumentFormat.OpenXml.Vml.ImageData>().FirstOrDefault();
                if (imageData?.RelationshipId?.Value is string vmlRelId)
                {
                    if (sourceMainPart.GetPartById(vmlRelId) is ImagePart sourceVmlImage)
                    {
                        var newVmlImagePart = targetMainPart.AddImagePart(sourceVmlImage.ContentType);
                        using var stream = sourceVmlImage.GetStream();
                        newVmlImagePart.FeedData(stream);
                        imageData.RelationshipId.Value = targetMainPart.GetIdOfPart(newVmlImagePart);
                    }
                }

                // Đối với line shapes và shapes khác, không cần xử lý relationship đặc biệt
                // Chúng sẽ được sao chép trực tiếp thông qua CloneNode(true)
            }

            // Xử lý VML Group nếu có
            foreach (var vmlGroup in element.Descendants<DocumentFormat.OpenXml.Vml.Group>())
            {
                // Groups có thể chứa nhiều shapes, xử lý đệ quy
                ProcessVmlElements(vmlGroup, sourceMainPart, targetMainPart);
            }
        }
    }
}
