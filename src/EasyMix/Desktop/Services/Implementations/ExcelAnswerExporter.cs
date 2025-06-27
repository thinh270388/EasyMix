using ClosedXML.Excel;
using Desktop.Models;
using Desktop.Models.Enums;
using Desktop.Services.Interfaces;

namespace Desktop.Services.Implementations
{
    public class ExcelAnswerExporter : IExcelAnswerExporter
    {
        public void ExportExcelAnswers(string filePath, List<QuestionExport> answers)
        {
            using var workbook = new XLWorkbook();
            var worksheet = workbook.Worksheets.Add("Đáp án");

            // ===== 1. Tạo tiêu đề cột =====
            worksheet.Cell(1, 1).Value = "Đề/Câu";

            // Câu trắc nghiệm 1-40
            for (int i = 1; i <= 40; i++)
                worksheet.Cell(1, i + 1).Value = i.ToString();

            // Câu đúng sai: 1a → 8d
            int col = 42;
            for (int i = 1; i <= 8; i++)
            {
                foreach (var opt in new[] { "a", "b", "c", "d" })
                {
                    worksheet.Cell(1, col++).Value = $"{i}{opt}";
                }
            }

            // Câu trả lời ngắn: 1-6
            for (int i = 1; i <= 6; i++)
                worksheet.Cell(1, col++).Value = $"{i}";

            // ===== 2. Ghi đáp án của tất cả mã đề =====
            var groupedAnswers = answers.GroupBy(a => a.Version).ToList();
            int row = 2;

            foreach (var group in groupedAnswers)
            {
                worksheet.Cell(row, 1).Value = group.Key; // Ghi mã đề

                // Khai báo danh sách đáp án
                var multipleChoiceAnswers = group
                    .Where(q => q.Type == QuestionType.MultipleChoice)
                    .OrderBy(q => q.QuestionNumber)
                    .Select(q => q.CorrectAnswer)
                    .ToList();

                var trueFalseAnswers = group
                    .Where(q => q.Type == QuestionType.TrueFalse)
                    .OrderBy(q => q.QuestionNumber)
                    .Select(q => q.CorrectAnswer)
                    .ToList();

                var shortAnswers = group
                    .Where(q => q.Type == QuestionType.ShortAnswer)
                    .OrderBy(q => q.QuestionNumber)
                    .Select(q => q.CorrectAnswer)
                    .ToList();

                // Trắc nghiệm: cột 2 → 41
                for (int i = 0; i < multipleChoiceAnswers.Count && i < 40; i++)
                    worksheet.Cell(row, i + 2).Value = multipleChoiceAnswers[i];

                // Đúng/Sai: cột 42 → 73 (8x4)
                for (int i = 0; i < trueFalseAnswers.Count && i < 32; i++)
                {
                    // Tách chuỗi đáp án
                    var answersArray = trueFalseAnswers[i]!.Split(new[] { ' ' }); // Tách ra thành 4 phần
                    for (int j = 0; j < 4; j++)
                    {
                        if (j < answersArray.Length)
                        {
                            worksheet.Cell(row, 42 + (i * 4) + j).Value = answersArray[(2 * j) + 1].Contains("Đúng") ? "Đ" : "S"; // Ghi "Đ" hoặc "S"
                        }
                    }
                }

                // Tự luận ngắn: cột 74 → 79
                for (int i = 0; i < shortAnswers.Count && i < 6; i++)
                    worksheet.Cell(row, i + 74).Value = shortAnswers[i];

                row++;
            }

            worksheet.Columns().AdjustToContents();
            workbook.SaveAs(filePath);
        }
    }
}
