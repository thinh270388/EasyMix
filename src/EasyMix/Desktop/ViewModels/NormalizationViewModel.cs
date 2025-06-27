using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Desktop.Helpers;
using Desktop.Models;
using Desktop.Models.Enums;
using Desktop.Services.Interfaces;
using Microsoft.Office.Interop.Word;
using MTGetEquationAddin;
using System.Collections.ObjectModel;
using System.IO;
using System.Reflection;
using System.Runtime.InteropServices;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;
using Task = System.Threading.Tasks.Task;
using Word = Microsoft.Office.Interop.Word;

namespace Desktop.ViewModels
{
    public partial class NormalizationViewModel : ObservableObject
    {
        private readonly IOpenXMLService _openXMLService;
        private readonly IInteropWordService _interopWordService;

        [ObservableProperty] private string sourceFile = string.Empty;
        [ObservableProperty] private string destinationFile = string.Empty;
        [ObservableProperty] private ObservableCollection<Question> questions = new ObservableCollection<Question>();
        [ObservableProperty] private bool isCenterImage = true;
        [ObservableProperty] private bool isBorderImage = true;

        public NormalizationViewModel(IOpenXMLService openXMLService, IInteropWordService interopWordService)
        {
            _openXMLService = openXMLService;
            _interopWordService = interopWordService;
        }

        [RelayCommand]
        private async Task BrowseFile()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    Title = "Chọn file Word"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    SourceFile = openFileDialog.FileName;
                    string folder = Path.Combine(Path.GetDirectoryName(SourceFile)!, Path.GetFileNameWithoutExtension(SourceFile));
                    if (!Directory.Exists(folder))
                        Directory.CreateDirectory(folder);
                    else
                    {
                        Directory.Delete(folder, true);
                        Directory.CreateDirectory(folder);
                    }

                    DestinationFile = $"{folder}\\EasyMix_{Path.GetFileName(SourceFile)}";
                    if (File.Exists(DestinationFile))
                        File.Delete(DestinationFile);
                    File.Copy(SourceFile, DestinationFile);

                    await ProcessDocumentAsync(DestinationFile, TypeExam.Normal);

                    var result = await _openXMLService.ParseDocxQuestionsAsync(DestinationFile);
                    Questions = new ObservableCollection<Question>(result);

                    MessageHelper.Success("Chuẩn hóa thành công!");
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        [RelayCommand]
        private async Task AnalyzeFile()
        {
            try
            {
                OpenFileDialog openFileDialog = new OpenFileDialog
                {
                    Filter = "Word Documents (*.docx)|*.docx",
                    Title = "Chọn file Word"
                };

                if (openFileDialog.ShowDialog() == true)
                {
                    SourceFile = openFileDialog.FileName;

                    var result = await _openXMLService.ParseDocxQuestionsAsync(SourceFile);
                    Questions = new ObservableCollection<Question>(result);
                }
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        [RelayCommand]
        private void OpenFile()
        {
            if (File.Exists(DestinationFile))
            {
                try
                {
                    System.Diagnostics.Process.Start(new System.Diagnostics.ProcessStartInfo
                    {
                        FileName = DestinationFile,
                        UseShellExecute = true
                    });
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }
            }
            else
            {
                MessageHelper.Error("Tệp không tồn tại");
            }
        }

        private async Task ProcessDocumentAsync(string filePath, TypeExam typeExam)
        {
            _Document? document = null;
            try
            {

                document = await _interopWordService.OpenDocumentAsync(filePath, visible: true);
                document.Activate();

                await _interopWordService.FormatDocumentAsync(document);
                await _interopWordService.DeleteAllHeadersAndFootersAsync(document);
                await _interopWordService.ConvertListFormatToTextAsync(document);

                var replacements = new Dictionary<string, string>
                {
                    ["^t"] = " ",
                    ["^l"] = " ",
                    ["^s"] = " ",
                    ["<$>"] = "^p<$>",
                    ["A. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["B. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["C. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["D. "] = "^p" + Constants.ANSWER_TEMPLATE,
                    ["<#>"] = Constants.QUESTION_TEMPLATE,
                    ["#"] = Constants.QUESTION_TEMPLATE,
                    ["[<br>]"] = Constants.QUESTION_TEMPLATE,
                    ["^p "] = "^p",
                    [" ^p"] = "^p",
                    ["  "] = " "
                };

                await _interopWordService.FindAndReplaceAsync(document, replacements, matchCase: true, matchWholeWord: false);

                await _interopWordService.FindAndReplaceRedToUnderlinedAsync(document);

                Word.Range range = document.Range();
                range.Font.Name = "Times New Roman";
                range.Font.Size = 12;
                range.Font.Color = WdColor.wdColorBlack;

                foreach (Word.Paragraph paragraph in document.Paragraphs)
                {
                    try
                    {
                        // Kiểm tra nếu đoạn có OLEObject (MathType)
                        if (paragraph.Range.InlineShapes.Count > 0)
                        {
                            foreach (Word.InlineShape shape in paragraph.Range.InlineShapes)
                            {
                                if (shape.Type == Word.WdInlineShapeType.wdInlineShapeEmbeddedOLEObject)
                                {
                                    if (shape.OLEFormat.ProgID == "Equation.3")
                                    {
                                        // Sửa căn chỉnh
                                        paragraph.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                        paragraph.SpaceBefore = 0;
                                        paragraph.SpaceAfter = 0;
                                        paragraph.LineSpacingRule = Word.WdLineSpacing.wdLineSpaceSingle;

                                        // Đôi khi cần reset vị trí baseline
                                        shape.Range.Font.Position = 0;
                                    }
                                }
                            }
                        }

                        paragraph.set_Style("Normal");
                        string str = paragraph.Range.Text.Trim();
                        string[] removeStarts = new[]
                        {
                            "phần 1", "phần 2", "phần 3", "phần 4",
                            "phần i", "phần ii", "phần iii", "phần iv",
                            "dạng 1", "dạng 2", "dạng 3", "dạng 4",
                            "dạng i", "dạng ii", "dạng iii", "dạng iv",
                            "i.", "ii.", "iii.", "iv.",
                            "<g0>", "<g1>", "<g2>", "<g3>",
                            "<#g0>", "<#g1>", "<#g2>", "<#g3>"
                        };
                        if (string.IsNullOrEmpty(str) || str.Equals(Constants.QUESTION_TEMPLATE) || str.Equals(Constants.ANSWER_TEMPLATE) ||
                            removeStarts.Any(prefix => str.StartsWith(prefix, StringComparison.OrdinalIgnoreCase)))
                        {
                            paragraph.Range.Delete();
                            continue;
                        }

                        await _interopWordService.ClearTabStopsAsync(paragraph);

                        // Chuẩn hóa định dạng "Câu ?"
                        if (str.StartsWith("câu "))
                        {
                            string[] patterns = { "Câu ? ", "Câu ?? ", "Câu ??? ", "Câu ?:", "Câu ??:", "Câu ???:", "Câu ?.", "Câu ??.", "Câu ???." };
                            foreach (var pattern in patterns)
                                await _interopWordService.FindAndReplaceFirstAsync(paragraph, pattern, Constants.QUESTION_TEMPLATE, matchWildcards: true);
                        }

                        // Chuẩn hóa a./b./c./d.
                        string[] keys = { "a.", "b.", "c.", "d.", "a)", "b)", "c)", "d)", " a.", " b.", " c.", " d.", " a)", " b)", " c)", " d)" };
                        foreach (var key in keys)
                        {
                            if (str.StartsWith(key))
                            {
                                string label = key.Trim().Substring(0, 1) + ") ";
                                await _interopWordService.FindAndReplaceFirstAsync(paragraph, key.Trim(), label);
                                break;
                            }
                        }

                        // Nếu chỉ chứa 1 hình ảnh và "/"
                        if (IsCenterImage && paragraph.Range.InlineShapes.Count == 1 && str == "/")
                        {
                            paragraph.Range.ParagraphFormat.Alignment = Word.WdParagraphAlignment.wdAlignParagraphCenter;
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(paragraph);
                    }
                }

                // Xử lý biểu tượng câu hỏi tùy theo loại đề
                string symbolQuestion = typeExam switch
                {
                    TypeExam.InTest => "<#>",
                    TypeExam.MCMix => "[<br>]",
                    TypeExam.SmartTest => "#",
                    _ => string.Empty
                };

                if (string.IsNullOrEmpty(symbolQuestion))
                    await _interopWordService.SetQuestionsToNumberAsync(document);
                else
                    await _interopWordService.FindAndReplaceAsync(document, new Dictionary<string, string> { [Constants.QUESTION_TEMPLATE] = symbolQuestion }, matchCase: true);

                // Xử lý đáp án
                if (typeExam == TypeExam.InTest)
                {
                    await _interopWordService.FindAndReplaceAsync(document, new Dictionary<string, string> { [Constants.ANSWER_TEMPLATE] = "<$>" }, matchCase: true);
                }
                else
                {
                    await _interopWordService.SetAnswersToABCDAsync(document);
                }

                await _interopWordService.FormatQuestionAndAnswerAsync(document);

                if (IsBorderImage)
                {
                    await _interopWordService.ProcessImagesInDocumentAsync(document, IsBorderImage);
                }

                await _interopWordService.FindAndReplaceAsync(document, new Dictionary<string, string>
                {
                    ["  "] = " ",
                    ["^p "] = "^p",
                    [" ^p"] = "^p"
                });

                foreach (Word.OMath math in document.OMaths)
                {
                    try
                    {
                        math.Range.Font.Name = "Cambria Math";
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(math);
                    }
                }

                foreach (Word.Field field in document.Fields)
                {
                    try
                    {
                        if (field.Code.Text.TrimStart().StartsWith("EQ ", StringComparison.OrdinalIgnoreCase))
                        {
                            field.Result.Font.Name = "Cambria Math";
                        }
                    }
                    finally
                    {
                        Marshal.ReleaseComObject(field);
                    }
                }

                try
                {
                   Connect connect = new Connect();
                    if (document != null)
                    {
                        Word.InlineShapes shapes = (Word.InlineShapes)document.GetType().InvokeMember("InlineShapes", BindingFlags.GetProperty, null, document, null)!;

                        int numShapesIterated = 0;

                        // Iterate over all of the shapes in the collection.
                        if (shapes != null && shapes.Count > 0)
                        {
                            numShapesIterated = connect.IterateShapes(ref shapes, true, true);
                        }
                    }
                }
                catch (Exception ex)
                {
                    MessageHelper.Error(ex);
                }

                await _interopWordService.SaveDocumentAsync(document!);
                MessageHelper.Success("Chuẩn hóa thành công!");
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
            finally
            {
                if (document != null)
                {
                    await _interopWordService.CloseDocumentAsync(document);
                    await _interopWordService.DisposeAsync();
                }
            }
        }
    }
}
