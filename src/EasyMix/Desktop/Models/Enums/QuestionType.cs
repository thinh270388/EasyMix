using System.ComponentModel;

namespace Desktop.Models.Enums
{
    public enum QuestionType
    {
        [Description("Trắc nghiệm đơn")]
        MultipleChoice,
        [Description("Đúng/sai")]
        TrueFalse,
        [Description("Trả lời ngắn")]
        ShortAnswer,
        [Description("Tự luận")]
        Essay,
        [Description("Chưa biết")]
        Unknown
    }
}
