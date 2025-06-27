using System.Text.RegularExpressions;

namespace Desktop.Helpers
{
    public class Constants
    {
        public const string QUESTION_TEMPLATE = "<CH>";
        public const string ANSWER_TEMPLATE = "<DA>";

        public const string ROOT_CODE = "000";

        public static readonly HashSet<string> QuestionPrefixes = new()
        {
            Constants.QUESTION_TEMPLATE, "<#>", "#", "[<br>]", "<G>", "<g>", "<NB>", "<TH>", "<VD>", "<VDC>"
        };

        public static readonly string[] AnswerPrefixes = { "A.", "B.", "C.", "D.", "A:", "B:", "C:", "D:", "a)", "b)", "c)", "d)", Constants.ANSWER_TEMPLATE, "<$>" };

        public const int TABSTOP_1 = 238;
        public const int TABSTOP_2 = 2619;
        public const int TABSTOP_3 = 5239;
        public const int TABSTOP_4 = 7859;

        public const string FONT_NAME = "Times New Roman";
        public const int FONT_SIZE = 12;

        // Hằng số cho các đường dẫn
        public const string TEMPLATES_FOLDER = "Assets\\Templates";
        public const string MIX_TEMPLATE_FILE = "TieuDe.docx";
        public const string GUIDE_TEMPLATE_FILE = "HuongDanGiai.docx";
        public const string ANSWERS_FOLER = "DapAn";
        public const string EXAM_PREFIX = "De_";
        public const string ANSWER_PREFIX = "DapAn_";
        public const string EXCEL_ANSWER_FILE = "DapAn.xlsx";

        public static readonly string[] ROMANS = { "I", "II", "III", "IV" };
        public static readonly string[] TITLES = { "PHẦN {0}. Câu hỏi trắc nghiệm nhiều lựa chọn. Thí sinh trả lời từ câu 1 đến câu {1}. Mỗi câu hỏi thí sinh chỉ chọn một phương án.",
                                   "PHẦN {0}. Câu hỏi trắc nghiệm đúng sai. Thí sinh trả lời từ câu 1 đến câu {1}. Trong mỗi ý a), b), c), d) ở mỗi câu, thí sinh chọn đúng hoặc sai.",
                                   "PHẦN {0}. Câu hỏi trắc nghiệm trả lời ngắn. Thí sinh trả lời từ câu 1 đến câu {1}.",
                                   "PHẦN {0}. Câu hỏi tự luận. Thí sinh trả lời từ câu 1 đến câu {1}." };


        public static readonly Regex QuestionHeaderRegex = new(@"^Câu\s+\d+[\.:]?", RegexOptions.IgnoreCase | RegexOptions.Compiled);
        public static readonly Regex MultipleChoiceAnswerRegex = new(@"^[A-Z]\.", RegexOptions.Compiled);
        public static readonly Regex TrueFalseAnswerRegex = new(@"^[a-d]\)", RegexOptions.Compiled);
        public static readonly Regex LevelRegex = new(@"\((NB|TH|VD)\)$", RegexOptions.IgnoreCase | RegexOptions.Compiled);
    }
}
