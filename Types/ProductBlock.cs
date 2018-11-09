using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelToXML.Types
{
    public class ProductBlock
    {
        /// <summary>
        ///     0 -- regular 
        ///     1 -- short
        /// </summary>
        public int productType { get; set; }
        public int worksheet { get; set; }
        public string worksheetName { get; set; }
        public string name { get; set; }
        public int start { get; set; }
        public int end { get; set; }
    }
}
