using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MyExcelMAUIApp.Models
{
    public class Cell
    {
        public string Expression { get; set; } = string.Empty;

        public object? Value { get; set; }

        public List<string> Dependencies { get; set; } = new List<string>();
    }
}
