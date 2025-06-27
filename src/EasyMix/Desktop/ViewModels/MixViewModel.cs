using CommunityToolkit.Mvvm.ComponentModel;
using CommunityToolkit.Mvvm.Input;
using Desktop.Helpers;
using Desktop.Models;
using Desktop.Services.Interfaces;
using DocumentFormat.OpenXml;
using Microsoft.Win32;
using System.Collections.ObjectModel;
using System.IO;
using System.Threading.Tasks;
using System.Windows.Documents;
using OpenFileDialog = Microsoft.Win32.OpenFileDialog;

namespace Desktop.ViewModels
{
    public partial class MixViewModel : ObservableObject
    {
        private readonly IOpenXMLService _openXMLService;
        private readonly IInteropWordService _interopWordService;

        private const string XmlFilePath = "config.xml";

        [ObservableProperty] private string outputFolder = string.Empty;
        [ObservableProperty] private string sourceFile = string.Empty;
        [ObservableProperty] private bool isEnableMix = false;
        [ObservableProperty] private MixInfo mixInfo = new();
        [ObservableProperty] private string examCodes = string.Empty;
        [ObservableProperty] private ObservableCollection<Question> questions = new ObservableCollection<Question>();
        [ObservableProperty] private bool isOK = false;
        [ObservableProperty] private FixedDocumentSequence? document;

        [ObservableProperty] private bool isCenterImage = true;
        [ObservableProperty] private bool isBorderImage = true;

        public MixViewModel(IOpenXMLService openXMLService, IInteropWordService interopWordService)
        {
            _openXMLService = openXMLService;
            _interopWordService = interopWordService;
            MixInfo = XmlHelper.LoadFromXml(XmlFilePath);
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
                    OutputFolder = Path.Combine(Path.GetDirectoryName(SourceFile)!, "EasyMix");
                    if (!Directory.Exists(OutputFolder))
                        Directory.CreateDirectory(OutputFolder);
                    else
                    {
                        Directory.Delete(OutputFolder, true);
                        Directory.CreateDirectory(OutputFolder);
                    }

                    var result = await _openXMLService.ParseDocxQuestionsAsync(SourceFile);
                    Questions = new ObservableCollection<Question>(result);

                    //string xpsPath = _interopWordService.ConvertDocxToXps(SourceFile);
                    //using var xpsDoc = new XpsDocument(xpsPath, FileAccess.Read);
                    //Document = xpsDoc.GetFixedDocumentSequence();

                    IsOK = Questions.All(q => q.IsValid);
                }
                IsEnableMix = !string.IsNullOrEmpty(SourceFile) && File.Exists(SourceFile) && IsOK;
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        [RelayCommand]
        private void Mix()
        {
            if (!File.Exists(SourceFile))
                return;
            try
            {
                var versions = ExamCodes.Split(' ', StringSplitOptions.RemoveEmptyEntries);
                if (versions.Count() == 0) return;

                MixInfo.Versions = versions;
                _openXMLService.GenerateShuffledExamsAsync(SourceFile, OutputFolder, MixInfo);

                MessageHelper.Success("Đã tạo xong file trộn đề và file đáp án");
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        [RelayCommand]
        private void SaveMixInfo()
        {
            try
            {
                XmlHelper.SaveToXml(XmlFilePath, MixInfo);
                MessageHelper.Success("Đã lưu thông tin cấu hình");
            }
            catch (Exception ex)
            {
                MessageHelper.Error(ex);
            }
        }

        [RelayCommand]
        private void GenerateRandomExamCodes()
        {
            var codes = new HashSet<string>();
            Random random = new Random();

            codes.Add("000");
            for (int i = 0; i < MixInfo.NumberOfVersions; i++)
            {
                string code = $"{(i % 9 + 1)}{random.Next(99):D2}"; // Tạo mã đề "1xx", "2xx", ..., "9xx"
                codes.Add(code);
            }

            ExamCodes = string.Join(" ", codes.OrderBy(c => c)); // Lưu mã đề vào biến ExamCodes
        }

        [RelayCommand]
        private void GenerateSequentialExamCodes()
        {
            var codes = new List<string> { "000" } // Thêm "000" làm phần tử đầu tiên
                .Concat(Enumerable.Range(0, MixInfo.NumberOfVersions)
                .Select(i => ((Convert.ToInt32(MixInfo.StartCode) * 100) + (i + 1)).ToString())); // Sinh mã đề liên tục bắt đầu từ StartCode

            ExamCodes = string.Join(" ", codes); // Lưu mã đề vào biến ExamCodes
        }
    }
}
