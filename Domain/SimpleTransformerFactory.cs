using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using XlsToXlsx.Interfaces;

namespace XlsToXlsx.Domain
{
    public class SimpleTransformerFactory
    {
        public Transformer CreateTransformer(string type)
        {
            Transformer transformer = null;

            if (type == "excel")
                transformer = new ExcelTransformer();
            else if (type == "word")
                transformer = new WordTransformer();

            return transformer;
        }
    }
}
