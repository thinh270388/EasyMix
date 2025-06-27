using System.ComponentModel;

namespace Desktop.Models.Enums
{
    public enum TypeExam
    {
        [Description("InTest")]
        InTest = 0,
        [Description("McMix")]
        MCMix = 1,
        [Description("Smart Test")]
        SmartTest = 2,
        [Description("Normal")]
        Normal = 3
    }
}
