using Desktop.Models.Enums;

namespace Desktop.Models
{
    public class Question
    {
        public QuestionType QuestionType { get; set; }
        public int Code { get; set; }
        public string? QuestionText { get; set; }
        public int CountAnswer { get; set; }
        public string? Answers { get; set; }
        public string? CorrectAnswer { get; set; }
        public string? Description { get; set; }
        public bool IsValid { get; set; } = false;
        public Level Level { get; set; } = Level.Know;
    }
}
