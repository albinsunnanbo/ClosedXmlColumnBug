using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestExcelExport
{
    public class DoesntWorkModel
    {
        [Display(Name = "First string")]
        public string String1 { get; set; }

        [Display(Name = "The Int")]
        public int? Int1 { get; set; }

        [Display(Name = "Second string")]
        public string String2 { get; set; }
    }
}
