using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToXML.Types
{
    public class ProductBlock
    {
        public int worksheet { get; set; }
        public string worksheetName { get; set; }
        public string name { get; set; }
        public int start { get; set; }
        public int end { get; set; }
    }
}
