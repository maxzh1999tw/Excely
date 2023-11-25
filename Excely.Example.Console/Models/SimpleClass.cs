using System.ComponentModel.DataAnnotations;

namespace Excely.Example.Console.Models
{
    internal class SimpleClass
    {
        [Display(Name = "序號")]
        public int Id { get; set; }

        [Display(Name = "文字欄位")]
        public string? StringField { get; set; }

        [Display(Name = "是/否")]
        public bool? BoolField { get; set; }

        [Display(Name = "日期欄位")]
        public DateTime? DateTimeField { get; set; }
    }
}
