using System.ComponentModel;

namespace Desktop.Models.Enums
{
    public enum Level
    {
        [Description("Nhận biết")]
        Know,
        [Description("Thông hiểu")]
        Understand,
        [Description("Vận dụng")]
        Manipulate,
        [Description("")]
        None
    }
}
